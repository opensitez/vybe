//! PHP-specific timestamp syntax layered over the shared date primitives.
use vybe_runtime::{Framework, Value};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:datetime",
        "microtime",
        Box::new(|_, args| {
            use std::time::{SystemTime, UNIX_EPOCH};

            let now = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default();
            let seconds = now.as_secs();
            let fraction = now.subsec_micros();
            if args.first().is_some_and(Value::as_bool) {
                Value::F64(seconds as f64 + fraction as f64 / 1_000_000.0)
            } else {
                Value::String(format!("0.{fraction:06} {seconds}").into())
            }
        }),
    );
    fw.register_host_fn(
        "php:datetime",
        "parseUnix",
        Box::new(|_, args| {
            let Some(Value::String(text)) = args.first() else {
                return Value::F64(f64::NAN);
            };
            let Some(number) = text.strip_prefix('@') else {
                return Value::F64(f64::NAN);
            };
            let digits = number
                .strip_prefix('-')
                .or_else(|| number.strip_prefix('+'))
                .unwrap_or(number);
            let mut parts = digits.split('.');
            let whole = parts.next().unwrap_or("");
            let fraction = parts.next();
            let valid = !whole.is_empty()
                && whole.bytes().all(|byte| byte.is_ascii_digit())
                && fraction.is_none_or(|part| {
                    !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit())
                })
                && parts.next().is_none();
            let seconds = if valid {
                number.parse::<f64>().unwrap_or(f64::NAN)
            } else {
                f64::NAN
            };
            Value::F64(seconds * 1000.0)
        }),
    );
}
