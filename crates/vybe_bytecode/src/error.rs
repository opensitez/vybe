use std::fmt;

/// VM runtime error with optional call stack trace.
#[derive(Debug, Clone)]
pub struct VMError {
    pub message: String,
    pub line: Option<u32>,
    /// Call stack at the point of error: (chunk_name, offset, line).
    /// Most recent frame first (like a stack trace).
    pub call_stack: Vec<StackFrame>,
}

/// A single frame in the error call stack.
#[derive(Debug, Clone)]
pub struct StackFrame {
    pub chunk_name: String,
    pub offset: usize,
    pub line: Option<u32>,
}

impl VMError {
    pub fn new(msg: impl Into<String>) -> Self {
        VMError {
            message: msg.into(),
            line: None,
            call_stack: Vec::new(),
        }
    }

    pub fn with_line(mut self, line: u32) -> Self {
        self.line = Some(line);
        self
    }

    pub fn with_stack(mut self, stack: Vec<StackFrame>) -> Self {
        self.call_stack = stack;
        self
    }
}

impl fmt::Display for VMError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "RuntimeError: {}", self.message)?;
        if let Some(line) = self.line {
            write!(f, " (line {})", line)?;
        }
        if !self.call_stack.is_empty() {
            write!(f, "\n  Call stack:")?;
            for frame in &self.call_stack {
                write!(f, "\n    at {} (offset {}", frame.chunk_name, frame.offset)?;
                if let Some(line) = frame.line {
                    write!(f, ", line {}", line)?;
                }
                write!(f, ")")?;
            }
        }
        Ok(())
    }
}

impl std::error::Error for VMError {}
