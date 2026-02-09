use crate::value::{RuntimeError, Value};
use std::fs;
use std::path::Path;

// ─── System.IO.File ───

pub fn file_readalltext_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("File.ReadAllText requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::read_to_string(&path) {
        Ok(content) => Ok(Value::String(content)),
        Err(e) => Err(RuntimeError::Custom(format!("File.ReadAllText error: {}", e))),
    }
}

pub fn file_writealltext_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("File.WriteAllText requires path and content arguments".to_string()));
    }
    let path = args[0].as_string();
    let content = args[1].as_string();
    match fs::write(&path, &content) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("File.WriteAllText error: {}", e))),
    }
}

pub fn file_appendalltext_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("File.AppendAllText requires path and content arguments".to_string()));
    }
    let path = args[0].as_string();
    let content = args[1].as_string();
    use std::io::Write;
    let mut file = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .map_err(|e| RuntimeError::Custom(format!("File.AppendAllText error: {}", e)))?;
    file.write_all(content.as_bytes())
        .map_err(|e| RuntimeError::Custom(format!("File.AppendAllText write error: {}", e)))?;
    Ok(Value::Nothing)
}

pub fn file_exists_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("File.Exists requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    Ok(Value::Boolean(Path::new(&path).is_file()))
}

pub fn file_delete_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("File.Delete requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::remove_file(&path) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("File.Delete error: {}", e))),
    }
}

pub fn file_copy_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("File.Copy requires source and destination arguments".to_string()));
    }
    let src = args[0].as_string();
    let dst = args[1].as_string();
    match fs::copy(&src, &dst) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("File.Copy error: {}", e))),
    }
}

pub fn file_move_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("File.Move requires source and destination arguments".to_string()));
    }
    let src = args[0].as_string();
    let dst = args[1].as_string();
    match fs::rename(&src, &dst) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("File.Move error: {}", e))),
    }
}

pub fn file_readalllines_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("File.ReadAllLines requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::read_to_string(&path) {
        Ok(content) => {
            let lines: Vec<Value> = content.lines().map(|l| Value::String(l.to_string())).collect();
            Ok(Value::Array(lines))
        }
        Err(e) => Err(RuntimeError::Custom(format!("File.ReadAllLines error: {}", e))),
    }
}

// ─── System.IO.Directory ───

pub fn directory_exists_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Directory.Exists requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    Ok(Value::Boolean(Path::new(&path).is_dir()))
}

pub fn directory_createdirectory_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Directory.CreateDirectory requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::create_dir_all(&path) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("Directory.CreateDirectory error: {}", e))),
    }
}

pub fn directory_delete_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Directory.Delete requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::remove_dir_all(&path) {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("Directory.Delete error: {}", e))),
    }
}

pub fn directory_getfiles_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Directory.GetFiles requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::read_dir(&path) {
        Ok(entries) => {
            let files: Vec<Value> = entries
                .filter_map(|e| e.ok())
                .filter(|e| e.path().is_file())
                .map(|e| Value::String(e.path().to_string_lossy().to_string()))
                .collect();
            Ok(Value::Array(files))
        }
        Err(e) => Err(RuntimeError::Custom(format!("Directory.GetFiles error: {}", e))),
    }
}

pub fn directory_getdirectories_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Directory.GetDirectories requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    match fs::read_dir(&path) {
        Ok(entries) => {
            let dirs: Vec<Value> = entries
                .filter_map(|e| e.ok())
                .filter(|e| e.path().is_dir())
                .map(|e| Value::String(e.path().to_string_lossy().to_string()))
                .collect();
            Ok(Value::Array(dirs))
        }
        Err(e) => Err(RuntimeError::Custom(format!("Directory.GetDirectories error: {}", e))),
    }
}

pub fn directory_getcurrentdirectory_fn() -> Result<Value, RuntimeError> {
    match std::env::current_dir() {
        Ok(path) => Ok(Value::String(path.to_string_lossy().to_string())),
        Err(e) => Err(RuntimeError::Custom(format!("Directory.GetCurrentDirectory error: {}", e))),
    }
}

// ─── System.IO.Path ───

pub fn path_combine_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Path.Combine requires two path arguments".to_string()));
    }
    let p1 = args[0].as_string();
    let p2 = args[1].as_string();
    let combined = Path::new(&p1).join(&p2);
    Ok(Value::String(combined.to_string_lossy().to_string()))
}

pub fn path_getfilename_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Path.GetFileName requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    let filename = Path::new(&path)
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();
    Ok(Value::String(filename))
}

pub fn path_getdirectoryname_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Path.GetDirectoryName requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    let dir = Path::new(&path)
        .parent()
        .map(|p| p.to_string_lossy().to_string())
        .unwrap_or_default();
    Ok(Value::String(dir))
}

pub fn path_getextension_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("Path.GetExtension requires a path argument".to_string()));
    }
    let path = args[0].as_string();
    let ext = Path::new(&path)
        .extension()
        .map(|e| format!(".{}", e.to_string_lossy()))
        .unwrap_or_default();
    Ok(Value::String(ext))
}

pub fn path_changeextension_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Path.ChangeExtension requires path and extension arguments".to_string()));
    }
    let path = args[0].as_string();
    let ext = args[1].as_string();
    let new_ext = ext.trim_start_matches('.');
    let new_path = Path::new(&path).with_extension(new_ext);
    Ok(Value::String(new_path.to_string_lossy().to_string()))
}
