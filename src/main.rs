mod com;
use com::*;

use windows::core;
use windows::Win32::System::Com::VARIANT;

fn main() -> core::Result<()> {
    init()?;

    let ws = ComObject::new("WScript.Shell")?;

    let cur_dir = ws.get_property("CurrentDirectory", None)?;
    println!("CurrentDirectory: {}", cur_dir.to_string()?);

    let popup_args = vec![
        VARIANT::from_str("ボタンを押してね"),
        VARIANT::from_i32(0),
        VARIANT::from_str("RustでCOMをやる"),
        VARIANT::from_i32(3),
    ];
    let pressed = ws.invoke_method("Popup", popup_args)?;
    println!("押されたボタン: {}", pressed.to_i32()?);

    uninit();
    Ok(())
}
