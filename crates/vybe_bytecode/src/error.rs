use std::fmt;

/// VM runtime error — language-agnostic.
#[derive(Debug, Clone)]
pub struct VMError {
    pub message: String,
    pub line: Option<u32>,
}

impl VMError {
    pub fn new(msg: impl Into<String>) -> Self {
        VMError { message: msg.into(), line: None }
    }

    pub fn with_line(mut self, line: u32) -> Self {
        self.line = Some(line);
        self
    }
}

impl fmt::Display for VMError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(line) = self.line {
            write!(f, "RuntimeError: {} (line {})", self.message, line)
        } else {
            write!(f, "RuntimeError: {}", self.message)
        }
    }
}

impl std::error::Error for VMError {}
