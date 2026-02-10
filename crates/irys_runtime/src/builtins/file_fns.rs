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

// ─── VB6-style File Functions ───

/// Dir([pathname[, attributes]]) - Returns file name matching pattern
pub fn dir_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    use std::cell::RefCell;
    
    thread_local! {
        static DIR_STATE: RefCell<Option<Vec<String>>> = RefCell::new(None);
        static DIR_INDEX: RefCell<usize> = RefCell::new(0);
    }
    
    if args.is_empty() {
        // Continue with previous search
        return DIR_STATE.with(|state| {
            DIR_INDEX.with(|index| {
                let mut idx = index.borrow_mut();
                let state_ref = state.borrow();
                
                if let Some(ref files) = *state_ref {
                    if *idx < files.len() {
                        let result = files[*idx].clone();
                        *idx += 1;
                        Ok(Value::String(result))
                    } else {
                        // Reset state
                        drop(state_ref);
                        *state.borrow_mut() = None;
                        *idx = 0;
                        Ok(Value::String(String::new()))
                    }
                } else {
                    Ok(Value::String(String::new()))
                }
            })
        });
    }
    
    let pattern = args[0].as_string();
    
    // Start new search
    DIR_INDEX.with(|index| *index.borrow_mut() = 0);
    
    let path = Path::new(&pattern);
    let (dir, pat) = if path.parent().is_some() && path.file_name().is_some() {
        (path.parent().unwrap(), path.file_name().unwrap().to_string_lossy().to_string())
    } else {
        (Path::new("."), pattern.clone())
    };
    
    match fs::read_dir(dir) {
        Ok(entries) => {
            let mut files = Vec::new();
            for entry in entries.flatten() {
                let filename = entry.file_name().to_string_lossy().to_string();
                if matches_pattern(&filename, &pat) {
                    files.push(filename);
                }
            }
            
            DIR_STATE.with(|state| {
                DIR_INDEX.with(|index| {
                    *state.borrow_mut() = Some(files.clone());
                    let mut idx = index.borrow_mut();
                    
                    if !files.is_empty() {
                        let result = files[0].clone();
                        *idx = 1;
                        Ok(Value::String(result))
                    } else {
                        Ok(Value::String(String::new()))
                    }
                })
            })
        }
        Err(e) => Err(RuntimeError::Custom(format!("Dir error: {}", e))),
    }
}

fn matches_pattern(filename: &str, pattern: &str) -> bool {
    // Simple wildcard matching (* and ?)
    let mut pat_chars = pattern.chars().peekable();
    let mut file_chars = filename.chars().peekable();
    
    loop {
        match (pat_chars.peek(), file_chars.peek()) {
            (None, None) => return true,
            (None, Some(_)) => return false,
            (Some(&'*'), _) => {
                pat_chars.next();
                if pat_chars.peek().is_none() {
                    return true;
                }
                // Try matching rest of pattern at each position
                let remaining_pat: String = pat_chars.collect();
                let remaining_file: String = file_chars.collect();
                for i in 0..=remaining_file.len() {
                    if matches_pattern(&remaining_file[i..], &remaining_pat) {
                        return true;
                    }
                }
                return false;
            }
            (Some(&'?'), Some(_)) => {
                pat_chars.next();
                file_chars.next();
            }
            (Some(&p), Some(&f)) if p.to_lowercase().eq(f.to_lowercase()) => {
                pat_chars.next();
                file_chars.next();
            }
            _ => return false,
        }
    }
}

/// FileCopy(source, destination) - Copies a file
pub fn filecopy_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("FileCopy requires 2 arguments".to_string()));
    }
    
    let source = args[0].as_string();
    let dest = args[1].as_string();
    
    fs::copy(&source, &dest)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("FileCopy error: {}", e)))
}

/// Kill(pathname) - Deletes a file
pub fn kill_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Kill requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    fs::remove_file(&path)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("Kill error: {}", e)))
}

/// Name(oldpathname, newpathname) - Renames a file or directory
pub fn name_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Name requires 2 arguments".to_string()));
    }
    
    let old_path = args[0].as_string();
    let new_path = args[1].as_string();
    
    fs::rename(&old_path, &new_path)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("Name error: {}", e)))
}

/// GetAttr(pathname) - Returns file attributes
pub fn getattr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("GetAttr requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    match fs::metadata(&path) {
        Ok(metadata) => {
            let mut attrs = 0;
            
            // vbNormal = 0, vbReadOnly = 1, vbHidden = 2, vbSystem = 4, 
            // vbDirectory = 16, vbArchive = 32
            if metadata.is_dir() {
                attrs |= 16;
            }
            
            #[cfg(unix)]
            {
                use std::os::unix::fs::PermissionsExt;
                let mode = metadata.permissions().mode();
                if mode & 0o200 == 0 {
                    attrs |= 1; // Read-only
                }
            }
            
            #[cfg(windows)]
            {
                use std::os::windows::fs::MetadataExt;
                attrs = metadata.file_attributes() as i32;
            }
            
            Ok(Value::Integer(attrs))
        }
        Err(e) => Err(RuntimeError::Custom(format!("GetAttr error: {}", e))),
    }
}

/// SetAttr(pathname, attributes) - Sets file attributes
pub fn setattr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("SetAttr requires 2 arguments".to_string()));
    }
    
    let path = args[0].as_string();
    let attrs = args[1].as_integer()?;
    
    // Simplified implementation - mainly handles read-only
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let metadata = fs::metadata(&path)
            .map_err(|e| RuntimeError::Custom(format!("SetAttr error: {}", e)))?;
        let mut permissions = metadata.permissions();
        
        if attrs & 1 != 0 {
            // Read-only
            permissions.set_mode(permissions.mode() & !0o200);
        } else {
            permissions.set_mode(permissions.mode() | 0o200);
        }
        
        fs::set_permissions(&path, permissions)
            .map(|_| Value::Nothing)
            .map_err(|e| RuntimeError::Custom(format!("SetAttr error: {}", e)))
    }
    
    #[cfg(windows)]
    {
        // Would use SetFileAttributes on Windows
        Ok(Value::Nothing)
    }
    
    #[cfg(not(any(unix, windows)))]
    {
        Ok(Value::Nothing)
    }
}

/// FileDateTime(pathname) - Returns file's last modified date/time
pub fn filedatetime_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("FileDateTime requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    match fs::metadata(&path) {
        Ok(metadata) => {
            let modified = metadata.modified()
                .map_err(|e| RuntimeError::Custom(format!("FileDateTime error: {}", e)))?;
            
            use chrono::{DateTime, Local};
            let dt: DateTime<Local> = DateTime::from(modified);
            Ok(Value::String(dt.format("%m/%d/%Y %H:%M:%S").to_string()))
        }
        Err(e) => Err(RuntimeError::Custom(format!("FileDateTime error: {}", e))),
    }
}

/// FileLen(pathname) - Returns file size in bytes
pub fn filelen_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("FileLen requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    match fs::metadata(&path) {
        Ok(metadata) => Ok(Value::Long(metadata.len() as i64)),
        Err(e) => Err(RuntimeError::Custom(format!("FileLen error: {}", e))),
    }
}

/// CurDir([drive]) - Returns current directory
pub fn curdir_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    match std::env::current_dir() {
        Ok(path) => Ok(Value::String(path.to_string_lossy().to_string())),
        Err(e) => Err(RuntimeError::Custom(format!("CurDir error: {}", e))),
    }
}

/// ChDir(path) - Changes current directory
pub fn chdir_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("ChDir requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    std::env::set_current_dir(&path)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("ChDir error: {}", e)))
}

/// MkDir(path) - Creates a directory
pub fn mkdir_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("MkDir requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    fs::create_dir(&path)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("MkDir error: {}", e)))
}

/// RmDir(path) - Removes a directory
pub fn rmdir_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("RmDir requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    fs::remove_dir(&path)
        .map(|_| Value::Nothing)
        .map_err(|e| RuntimeError::Custom(format!("RmDir error: {}", e)))
}

/// FreeFile() - Returns next available file number
pub fn freefile_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    use std::sync::atomic::{AtomicI32, Ordering};
    
    static FILE_NUMBER: AtomicI32 = AtomicI32::new(1);
    
    let num = FILE_NUMBER.fetch_add(1, Ordering::SeqCst);
    Ok(Value::Integer(num))
}

// ─── VB6 File Handle Functions ───
// Note: These work with the file_handles HashMap in the Interpreter
// For now, we provide basic implementations that work with file metadata

/// EOF(filenumber) - Returns True if at end of file
/// Note: In full implementation, this would check the file handle position
pub fn eof_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("EOF requires 1 argument".to_string()));
    }
    
    let _file_num = args[0].as_integer()?;
    
    // Without actual file handle tracking, we return False
    // A full implementation would track open file handles and their positions
    Ok(Value::Boolean(false))
}

/// LOF(filenumber) - Returns length of open file in bytes
pub fn lof_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("LOF requires 1 argument".to_string()));
    }
    
    let _file_num = args[0].as_integer()?;
    
    // Without actual file handle tracking, return 0
    // A full implementation would look up the file handle and get its length
    Ok(Value::Long(0))
}

/// LOC(filenumber) - Returns current read/write position in file
pub fn loc_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("LOC requires 1 argument".to_string()));
    }
    
    let _file_num = args[0].as_integer()?;
    
    // Without actual file handle tracking, return 0
    // A full implementation would track file position
    Ok(Value::Long(0))
}

/// FileAttr(filenumber, returntype) - Returns mode or OS file handle of open file
pub fn fileattr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(RuntimeError::Custom("FileAttr requires 1 or 2 arguments".to_string()));
    }
    
    let _file_num = args[0].as_integer()?;
    let return_type = if args.len() == 2 {
        args[1].as_integer()?
    } else {
        1 // Default to file mode
    };
    
    // returntype: 1 = file mode, 2 = OS file handle
    // Without actual file handle tracking, return default values
    match return_type {
        1 => Ok(Value::Integer(1)), // Input mode
        2 => Ok(Value::Integer(0)), // No OS handle
        _ => Err(RuntimeError::Custom("FileAttr: invalid return type".to_string())),
    }
}

// ─── Image Functions ───

/// LoadPicture(filename) - Loads an image file
/// Returns a simple object representing the image
pub fn loadpicture_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 1 {
        return Err(RuntimeError::Custom("LoadPicture requires 1 argument".to_string()));
    }
    
    let path = args[0].as_string();
    
    // Check if file exists
    if !Path::new(&path).exists() {
        return Err(RuntimeError::Custom(format!("LoadPicture: file not found: {}", path)));
    }
    
    // Read file metadata
    match fs::metadata(&path) {
        Ok(metadata) => {
            use std::rc::Rc;
            use std::cell::RefCell;
            use std::collections::HashMap;
            use crate::ObjectData;
            
            // Create a simple Picture object with basic properties
            let mut fields = HashMap::new();
            fields.insert("filename".to_string(), Value::String(path.clone()));
            fields.insert("size".to_string(), Value::Long(metadata.len() as i64));
            fields.insert("type".to_string(), Value::String("Picture".to_string()));
            
            // Determine image type from extension
            let ext = Path::new(&path)
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_lowercase();
            fields.insert("format".to_string(), Value::String(ext));
            
            let obj = ObjectData {
                class_name: "Picture".to_string(),
                fields,
            };
            
            Ok(Value::Object(Rc::new(RefCell::new(obj))))
        }
        Err(e) => Err(RuntimeError::Custom(format!("LoadPicture error: {}", e))),
    }
}

/// SavePicture(picture, filename) - Saves a picture object to file
pub fn savepicture_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("SavePicture requires 2 arguments".to_string()));
    }
    
    let _picture = &args[0]; // Picture object
    let dest_path = args[1].as_string();
    
    // In a full implementation, this would:
    // 1. Extract image data from picture object
    // 2. Encode it to the target format
    // 3. Write to file
    
    // For now, if the picture has a filename, copy that file
    if let Value::Object(obj_ref) = _picture {
        let obj = obj_ref.borrow();
        if let Some(Value::String(source)) = obj.fields.get("filename") {
            // Copy source file to destination
            match fs::copy(source, &dest_path) {
                Ok(_) => return Ok(Value::Nothing),
                Err(e) => return Err(RuntimeError::Custom(format!("SavePicture error: {}", e))),
            }
        }
    }
    
    // If we can't extract source, just create an empty file
    match fs::write(&dest_path, b"") {
        Ok(_) => Ok(Value::Nothing),
        Err(e) => Err(RuntimeError::Custom(format!("SavePicture error: {}", e))),
    }
}
