use std::io;

#[derive(Debug, thiserror::Error)]
pub enum SaveError {
    #[error("IO error: {0}")]
    Io(#[from] io::Error),
    #[error("Formatting error: {0}")]
    Fmt(#[from] std::fmt::Error),
    #[error("Parse error: {0}")]
    Parse(String),
}

pub type SaveResult<T> = Result<T, SaveError>;
