use crate::value::{RuntimeError, Value};

pub fn msgbox(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("MsgBox requires at least one argument".to_string()));
    }

    let message = args[0].as_string();

    // For now, just print to console
    // In a real implementation, this would show a GUI dialog
    println!("MsgBox: {}", message);

    Ok(Value::Integer(1)) // vbOK
}
