use windows::{
    core::{self, BSTR, HSTRING, ComInterface, GUID, PCWSTR},
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::{
                CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED,
                CLSIDFromString, CoCreateInstance, CLSCTX_ALL, CLSCTX_LOCAL_SERVER,
                IDispatch, DISPPARAMS,
                DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
                VARIANT, VARIANT_0_0,
                VARENUM, VT_NULL, VT_BSTR, VT_BOOL, VT_I4, VT_ARRAY, VT_BYREF, VT_VARIANT,
                SAFEARRAY,
            },
            Ole::{
                VariantClear, VariantChangeType, GetActiveObject,
                DISPID_PROPERTYPUT,
            },
        }
    }
};

use std::mem::ManuallyDrop;

const LOCALE_USER_DEFAULT: u32 = 0x400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

pub fn init() -> core::Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
    }
}
pub fn uninit() {
    unsafe {
        CoUninitialize();
    }
}
pub struct ComObject {
    disp: IDispatch,
}

#[allow(unused)]
impl ComObject {
    /// COMオブジェクトを新規に作成します
    ///
    /// ProgIDかCLSID文字列 ( {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} 形式) を渡す
    pub fn new(id: &str) -> core::Result<Self> {
        unsafe {
            let lpsz = HSTRING::from(id);
            let rclsid = CLSIDFromString(&lpsz)?;
            let disp = match CoCreateInstance(&rclsid, None, CLSCTX_ALL) {
                Ok(disp) => disp,
                Err(_) => CoCreateInstance(&rclsid, None, CLSCTX_LOCAL_SERVER)?,
            };
            Ok(Self { disp })
        }
    }
    /// 起動中のExcelを捕まえるなどで使う
    ///
    /// ProgIDかCLSID文字列を渡す
    pub fn get(id: &str) -> core::Result<Option<Self>> {
        unsafe {
            let lpsz = HSTRING::from(id);
            let rclsid = CLSIDFromString(&lpsz)?;
            let pvreserved = std::ptr::null_mut() as *mut std::ffi::c_void;
            let mut ppunk = None;
            GetActiveObject(&rclsid, pvreserved, &mut ppunk)?;
            let disp: Option<IDispatch> = match ppunk {
                Some(unk) => {
                    // ComInterfaceをuseしておくことでcastが使える
                    let disp = unk.cast()?;
                    Some(disp)
                },
                None => None,
            };
            Ok(disp.map(|disp| Self {disp}))
        }
    }
    fn get_id_from_name(&self, name: &str) -> core::Result<i32> {
        unsafe {
            let hstring = HSTRING::from(name);
            let rgsznames = PCWSTR::from_raw(hstring.as_ptr());
            let mut rgdispid = 0;
            self.disp.GetIDsOfNames(&GUID::zeroed(), &rgsznames, 1, LOCALE_USER_DEFAULT, &mut rgdispid)?;
            Ok(rgdispid)
        }
    }
    fn invoke(&self, dispidmember: i32, pdispparams: &DISPPARAMS, wflags: DISPATCH_FLAGS) -> core::Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::default();
            self.disp.Invoke(
                dispidmember,
                &GUID::zeroed(),
                LOCALE_SYSTEM_DEFAULT,
                wflags,
                pdispparams,
                Some(&mut result),
                None,
                None
            )?;
            Ok(result)
        }
    }
    /// プロパティの値を得ます
    ///
    /// 値を得たいプロパティの名前を渡してください
    /// パラメータ付きプロパティの場合はパラメータを示すVARIANTを渡します
    pub fn get_property(&self, prop: &str, param: Option<VARIANT>) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(prop)?;
        let mut pdispparams = DISPPARAMS::default();
        let mut args = if let Some(param) = param {
            vec![param]
        } else {
            vec![]
        };
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_PROPERTYGET)
    }
    /// プロパティに値をセットします
    ///
    /// プロパティ名と必要ならばそのパラメータとセットする値を渡します
    pub fn set_property(&self, prop: &str, param: Option<VARIANT>, value: VARIANT) -> core::Result<()> {
        let dispidmember = self.get_id_from_name(prop)?;
        let mut pdispparams = DISPPARAMS::default();
        let mut args = if let Some(param) = param {
            vec![param, value]
        } else {
            vec![value]
        };
        let mut named_args = vec![DISPID_PROPERTYPUT];
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        pdispparams.cNamedArgs = 1;
        pdispparams.rgdispidNamedArgs = named_args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_PROPERTYPUT)?;
        Ok(())
    }
    /// メソッドを実行します
    ///
    /// メソッド名とメソッドに渡す引数を渡します
    pub fn invoke_method(&self, method: &str, mut args: Vec<VARIANT>) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(method)?;
        let mut pdispparams = DISPPARAMS::default();
        args.reverse();
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_METHOD)
    }
}

pub trait VariantExt {
    /// VT_NULLなVARIANTを作る
    fn null() -> VARIANT;
    /// VT_BYREF|VT_VARIANTなVARIANTを作る、参照渡し用
    /// 引数は参照先となるVARIANT
    fn by_ref(var_val: *mut VARIANT) -> VARIANT;
    /// VT_I4を作る
    fn from_i32(n: i32) -> VARIANT;
    /// VT_BSTRを作る
    fn from_str(s: &str) -> VARIANT;
    /// VT_BOOLを作る
    fn from_bool(b: bool) -> VARIANT;
    /// VT_ARRAY|VT_VARIANTを作る
    fn from_safearray(psa: *mut SAFEARRAY) -> VARIANT;
    /// VARIANTをi32にする
    fn to_i32(&self) -> core::Result<i32>;
    /// VARIANTをStringにする
    fn to_string(&self) -> core::Result<String>;
    /// VARIANTをboolにする
    fn to_bool(&self) -> core::Result<bool>;
}

impl VariantExt for VARIANT {
    fn null() -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_NULL;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn by_ref(var_val: *mut VARIANT) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VARENUM(VT_BYREF.0|VT_VARIANT.0);
        v00.Anonymous.pvarVal = var_val;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i32(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_I4;
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_str(s: &str) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_BSTR;
        let bstr = BSTR::from(s);
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_bool(b: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_BOOL;
        v00.Anonymous.boolVal = VARIANT_BOOL::from(b);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_safearray(psa: *mut SAFEARRAY) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VARENUM(VT_ARRAY.0|VT_VARIANT.0);
        v00.Anonymous.parray = psa;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn to_i32(&self) -> core::Result<i32> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.lVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_string(&self) -> core::Result<String> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_BSTR)?;
            let v00 = &new.Anonymous.Anonymous;
            let str = v00.Anonymous.bstrVal.to_string();
            VariantClear(&mut new)?;
            Ok(str)
        }
    }
    fn to_bool(&self) -> core::Result<bool> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_BOOL)?;
            let v00 = &new.Anonymous.Anonymous;
            let b = v00.Anonymous.boolVal.as_bool();
            VariantClear(&mut new)?;
            Ok(b)
        }
    }
}