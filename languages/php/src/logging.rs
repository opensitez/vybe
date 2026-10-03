//! PHP error logging keeps its return value and destination semantics separate
//! from the WASI console interfaces.
use std::io::Write;
use vybe_runtime::{Framework, Value};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:logging",
        "errorLog",
        Box::new(|_, args| {
            let message = match args.first() {
                Some(Value::String(value)) => value.to_string(),
                Some(value) => value.to_string(),
                None => return Value::Bool(false),
            };
            let message_type = args.get(1).map(Value::as_i32).unwrap_or(0);
            match message_type {
                0 => {
                    eprintln!("{message}");
                    Value::Bool(true)
                }
                3 => {
                    let Some(Value::String(destination)) = args.get(2) else {
                        return Value::Bool(false);
                    };
                    let written = std::fs::OpenOptions::new()
                        .create(true)
                        .append(true)
                        .open(destination.as_ref())
                        .and_then(|mut file| file.write_all(message.as_bytes()))
                        .is_ok();
                    Value::Bool(written)
                }
                _ => Value::Bool(false),
            }
        }),
    );
}
