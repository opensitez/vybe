use crate::form::Form;
use std::fs;
use std::path::Path;

#[derive(Debug, thiserror::Error)]
pub enum SerializationError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
}

pub type SerializationResult<T> = Result<T, SerializationError>;

pub fn save_form(form: &Form, path: impl AsRef<Path>) -> SerializationResult<()> {
    let json = serde_json::to_string_pretty(form)?;
    fs::write(path, json)?;
    Ok(())
}

pub fn load_form(path: impl AsRef<Path>) -> SerializationResult<Form> {
    let json = fs::read_to_string(path)?;
    let form = serde_json::from_str(&json)?;
    Ok(form)
}
