//! Cryptographic random bytes for PHP. Each byte is stored as one Latin-1
//! code unit because the VM currently represents all strings as UTF-8 text.
//! This preserves byte values for PHP's `ord` and `bin2hex` operations.
use std::sync::Arc;
use vybe_runtime::{Framework, Value};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:random",
        "bytes",
        Box::new(|ctx, args| {
            let length = args.first().map_or(0, |value| value.as_f64() as i64);
            if length <= 0 {
                ctx.throw_value(Value::String(Arc::from(
                    "random_bytes(): Argument #1 ($length) must be greater than 0",
                )));
                return Value::Null;
            }
            let Ok(length) = usize::try_from(length) else {
                ctx.throw_value(Value::String(Arc::from("random_bytes(): invalid length")));
                return Value::Null;
            };
            let mut bytes = Vec::new();
            if bytes.try_reserve_exact(length).is_err() {
                ctx.throw_value(Value::String(Arc::from(
                    "random_bytes(): allocation failed",
                )));
                return Value::Null;
            }
            bytes.resize(length, 0);
            if getrandom::getrandom(&mut bytes).is_err() {
                ctx.throw_value(Value::String(Arc::from(
                    "random_bytes(): secure random source failed",
                )));
                return Value::Null;
            }
            Value::String(bytes.into_iter().map(char::from).collect::<String>().into())
        }),
    );
}
