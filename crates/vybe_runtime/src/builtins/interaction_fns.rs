use crate::value::{RuntimeError, Value};
use std::process::Command;
use std::collections::HashMap;
use std::rc::Rc;
use std::cell::RefCell;

/// Beep() - Makes a beep sound
pub fn beep_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    #[cfg(target_os = "macos")]
    {
        // Use afplay to play system beep
        let _ = Command::new("afplay")
            .arg("/System/Library/Sounds/Ping.aiff")
            .spawn();
    }
    
    #[cfg(target_os = "linux")]
    {
        // Use beep command or fall back to terminal bell
        if Command::new("beep").spawn().is_err() {
            print!("\x07"); // Terminal bell
        }
    }
    
    #[cfg(target_os = "windows")]
    {
        // Use Windows beep API via PowerShell
        let _ = Command::new("powershell")
            .arg("-Command")
            .arg("[console]::beep(800, 200)")
            .spawn();
    }
    
    Ok(Value::Nothing)
}

/// Shell(pathname[, windowstyle]) - Executes an external program
pub fn shell_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("Shell requires 1 or 2 arguments".to_string()));
    }
    
    let command = args[0].as_string();
    let _window_style = if args.len() == 2 {
        args[1].as_integer().unwrap_or(1)
    } else {
        1 // vbNormalFocus
    };
    
    // Parse command and arguments
    let parts: Vec<&str> = command.split_whitespace().collect();
    if parts.is_empty() {
        return Err(RuntimeError::Custom("Shell: empty command".to_string()));
    }
    
    let program = parts[0];
    let args = &parts[1..];
    
    match Command::new(program).args(args).spawn() {
        Ok(child) => {
            // Return process ID
            Ok(Value::Integer(child.id() as i32))
        }
        Err(e) => Err(RuntimeError::Custom(format!("Shell error: {}", e))),
    }
}

/// Environ(envstring | number) - Returns environment variable value
pub fn environ_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Environ requires 1 argument".to_string()));
    }
    
    match &args[0] {
        Value::String(var_name) => {
            // Get specific environment variable
            match std::env::var(var_name) {
                Ok(value) => Ok(Value::String(value)),
                Err(_) => Ok(Value::String(String::new())),
            }
        }
        Value::Integer(index) => {
            // Get environment variable by index
            let vars: Vec<(String, String)> = std::env::vars().collect();
            if *index >= 1 && (*index as usize) <= vars.len() {
                let (key, value) = &vars[(*index - 1) as usize];
                Ok(Value::String(format!("{}={}", key, value)))
            } else {
                Ok(Value::String(String::new()))
            }
        }
        _ => Err(RuntimeError::Custom("Environ requires string or integer argument".to_string())),
    }
}

/// Command() - Returns command line arguments
pub fn command_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    Ok(Value::String(args.join(" ")))
}

/// SendKeys(keys[, wait]) - Sends keystrokes to the active window (stub - platform specific)
pub fn sendkeys_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("SendKeys requires 1 or 2 arguments".to_string()));
    }
    
    let _keys = args[0].as_string();
    let _wait = if args.len() == 2 {
        args[1].as_bool().unwrap_or(false)
    } else {
        false
    };
    
    // SendKeys is highly platform-specific and requires accessibility permissions
    // For now, just return Nothing (real implementation would need platform-specific code)
    Ok(Value::Nothing)
}

/// AppActivate(title | pid) - Activates an application window (stub - platform specific)
pub fn appactivate_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("AppActivate requires 1 or 2 arguments".to_string()));
    }
    
    // AppActivate is platform-specific and requires window management APIs
    // Real implementation would use platform-specific code
    Ok(Value::Nothing)
}

/// Load form_object - Loads a form into memory (creates instance)
pub fn load_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Load requires exactly one argument".to_string()));
    }
    
    // In a full implementation, this would:
    // 1. Create a new form instance
    // 2. Initialize form controls
    // 3. Trigger Form_Load event
    // For now, just return Nothing
    Ok(Value::Nothing)
}

/// Unload form_object - Unloads a form from memory (destroys instance)
pub fn unload_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Unload requires exactly one argument".to_string()));
    }
    
    // In a full implementation, this would:
    // 1. Trigger Form_Unload event
    // 2. Destroy form controls
    // 3. Release form instance
    // For now, just return Nothing
    Ok(Value::Nothing)
}

/// App object - Returns application-level properties
pub fn app_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    use crate::ObjectData;
    
    let mut fields = HashMap::new();
    fields.insert("path".to_string(), Value::String(".".to_string()));
    fields.insert("title".to_string(), Value::String("vybe Application".to_string()));
    fields.insert("exename".to_string(), Value::String("vybe_app".to_string()));
    fields.insert("major".to_string(), Value::Integer(1));
    fields.insert("minor".to_string(), Value::Integer(0));
    fields.insert("revision".to_string(), Value::Integer(0));
    
    let obj = ObjectData {
        class_name: "App".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}

/// Screen object - Returns screen/display properties
pub fn screen_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    use crate::ObjectData;
    
    let mut fields = HashMap::new();
    // Default screen dimensions (can be platform-specific in full implementation)
    fields.insert("width".to_string(), Value::Integer(1920));
    fields.insert("height".to_string(), Value::Integer(1080));
    fields.insert("twipsperpixelx".to_string(), Value::Integer(15));
    fields.insert("twipsperpixely".to_string(), Value::Integer(15));
    
    let obj = ObjectData {
        class_name: "Screen".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}

/// Clipboard object - Returns clipboard object with methods
pub fn clipboard_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    use crate::ObjectData;
    
    let mut fields = HashMap::new();
    // Clipboard data stored as string (simplified)
    fields.insert("text".to_string(), Value::String(String::new()));
    
    let obj = ObjectData {
        class_name: "Clipboard".to_string(),
        fields,
        drawing_commands: Vec::new(),
    };
    
    Ok(Value::Object(Rc::new(RefCell::new(obj))))
}

/// Clipboard.GetText() - Get text from clipboard
pub fn clipboard_gettext_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    // In a full implementation, this would read from system clipboard
    // For now, return empty string
    Ok(Value::String(String::new()))
}

/// Clipboard.SetText(text) - Set text to clipboard
pub fn clipboard_settext_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Clipboard.SetText requires 1 argument".to_string()));
    }
    // In a full implementation, this would write to system clipboard
    // For now, just return Nothing
    Ok(Value::Nothing)
}

/// Clipboard.Clear() - Clear clipboard
pub fn clipboard_clear_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    // In a full implementation, this would clear system clipboard
    Ok(Value::Nothing)
}

/// Forms collection - Returns collection of all forms
pub fn forms_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    use crate::collections::ArrayList;
    // Return an empty collection for now
    // In a full implementation, this would return all loaded forms
    Ok(Value::Collection(Rc::new(RefCell::new(ArrayList::new()))))
}
