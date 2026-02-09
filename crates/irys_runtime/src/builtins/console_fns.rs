use crate::value::Value;

/// Join args into a single output string (Console.Write/WriteLine)
pub fn console_write_fn(args: &[Value]) -> String {
    args.iter().map(|v| v.as_string()).collect::<Vec<_>>().join("")
}

/// Stub for Console.ReadLine (no stdin in GUI context)
pub fn console_readline_fn() -> Value {
    Value::String(String::new())
}
