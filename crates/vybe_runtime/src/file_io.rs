use std::fs::{File, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::Path;
use crate::value::RuntimeError;

#[derive(Debug)]
pub enum FileHandle {
    Input(BufReader<File>),
    Output(BufWriter<File>),
    Append(BufWriter<File>),
    Binary(File),
}

pub fn open_file(path: &str, mode: vybe_parser::ast::stmt::FileOpenMode) -> Result<FileHandle, RuntimeError> {
    let file_path = Path::new(path);
    
    match mode {
        vybe_parser::ast::stmt::FileOpenMode::Input => {
            let file = File::open(file_path).map_err(|e| RuntimeError::Custom(format!("Failed to open file for Input: {}", e)))?;
            Ok(FileHandle::Input(BufReader::new(file)))
        }
        vybe_parser::ast::stmt::FileOpenMode::Output => {
            let file = File::create(file_path).map_err(|e| RuntimeError::Custom(format!("Failed to open file for Output: {}", e)))?;
            Ok(FileHandle::Output(BufWriter::new(file)))
        }
        vybe_parser::ast::stmt::FileOpenMode::Append => {
            let file = OpenOptions::new()
                .write(true)
                .append(true)
                .create(true)
                .open(file_path)
                .map_err(|e| RuntimeError::Custom(format!("Failed to open file for Append: {}", e)))?;
            Ok(FileHandle::Append(BufWriter::new(file)))
        }
        vybe_parser::ast::stmt::FileOpenMode::Binary => {
            let file = OpenOptions::new()
                .read(true)
                .write(true)
                .create(true)
                .open(file_path)
                .map_err(|e| RuntimeError::Custom(format!("Failed to open file for Binary: {}", e)))?;
            Ok(FileHandle::Binary(file))
        }
        vybe_parser::ast::stmt::FileOpenMode::Random => {
             // For now, treat Random as Binary or Error? Let's use Binary for flexibility
             let file = OpenOptions::new()
                .read(true)
                .write(true)
                .create(true)
                .open(file_path)
                .map_err(|e| RuntimeError::Custom(format!("Failed to open file for Random: {}", e)))?;
            Ok(FileHandle::Binary(file))
        }
    }
}

pub fn write_line(handle: &mut FileHandle, text: &str) -> Result<(), RuntimeError> {
    match handle {
        FileHandle::Output(writer) => {
            writeln!(writer, "{}", text).map_err(|e| RuntimeError::Custom(format!("Write error: {}", e)))
        }
        FileHandle::Append(writer) => {
            writeln!(writer, "{}", text).map_err(|e| RuntimeError::Custom(format!("Write error: {}", e)))
        }
        _ => Err(RuntimeError::Custom("File mode not valid for Output".to_string())),
    }
}

pub fn print_string(handle: &mut FileHandle, text: &str) -> Result<(), RuntimeError> {
    match handle {
        FileHandle::Output(writer) => {
            write!(writer, "{}", text).map_err(|e| RuntimeError::Custom(format!("Write error: {}", e)))
        }
        FileHandle::Append(writer) => {
            write!(writer, "{}", text).map_err(|e| RuntimeError::Custom(format!("Write error: {}", e)))
        }
        _ => Err(RuntimeError::Custom("File mode not valid for Output".to_string())),
    }
}


pub fn read_line(handle: &mut FileHandle) -> Result<String, RuntimeError> {
    match handle {
        FileHandle::Input(reader) => {
            let mut line = String::new();
            let bytes = reader.read_line(&mut line).map_err(|e| RuntimeError::Custom(format!("Read error: {}", e)))?;
            if bytes == 0 {
                return Err(RuntimeError::Custom("End of file encountered".to_string()));
            }
            // Trim trailing newline (created by read_line including \n or \r\n)
            // But Line Input usually returns the content without newline? 
            // VB6 Line Input reads a line and assigns it to a variable, excluding the end-of-line characters.
            let trimmed = line.trim_end_matches(|c| c == '\r' || c == '\n').to_string();
            Ok(trimmed)
        }
        _ => Err(RuntimeError::Custom("File mode not valid for Input".to_string())),
    }
}
