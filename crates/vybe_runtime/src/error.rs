use std::fmt;

/// VM runtime error with optional call stack trace.
#[derive(Debug, Clone)]
pub struct VMError {
    pub message: String,
    pub line: Option<u32>,
    /// Call stack at the point of error: (chunk_name, offset, line).
    /// Most recent frame first (like a stack trace).
    pub call_stack: Vec<StackFrame>,
    /// Optional, bounded debugger context for an unhandled failure.
    pub diagnostic: Option<String>,
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
            diagnostic: None,
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

    pub fn with_diagnostic(mut self, diagnostic: String) -> Self {
        self.diagnostic = Some(diagnostic);
        self
    }

    /// Is this a WASM TRAP (as opposed to a language-level exception that
    /// found no handler)? The distinction decides who may catch it:
    ///   * WASM 3.0 core — a trap is NOT catchable by `try_table`, not even
    ///     by `catch_all`.
    ///   * WebAssembly JS Interface — a trap crossing into host code IS
    ///     catchable there, surfacing as a `WebAssembly.RuntimeError`.
    ///
    /// `TRAP_PREFIX` is the single classifier, and it is not a new convention:
    /// every trap site in the VM already spells its message this way. Keep new
    /// trap messages prefixed or they become uncatchable at the host boundary.
    pub fn is_trap(&self) -> bool {
        self.message.starts_with(TRAP_PREFIX)
    }

    pub(crate) fn is_suspension(&self) -> bool {
        ["__await__:", "__jspi__:", "__future__:", "__stream_read__:"]
            .iter().any(|prefix| self.message.starts_with(prefix))
    }
}

/// The canonical prefix every VM trap message carries. Sole input to
/// [`VMError::is_trap`].
pub const TRAP_PREFIX: &str = "trap: ";

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
        if let Some(diagnostic) = &self.diagnostic {
            write!(f, "\n{diagnostic}")?;
        }
        Ok(())
    }
}

impl std::error::Error for VMError {}
