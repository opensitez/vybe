use std::sync::{Arc, Mutex};
use vybe_runtime::{VM, Value};

pub fn register(vm: &mut VM, output: Arc<Mutex<Vec<String>>>) {
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(move |ctx, args| {
            let (future, id) = ctx.create_future();
            let output = Arc::clone(&output);
            let mut received = Vec::new();
            let result = args.first().ok_or("invalid stream").and_then(|stream| {
                ctx.consume_byte_stream(
                    stream,
                    id,
                    Box::new(move |bytes| {
                        if let Some(bytes) = bytes {
                            received.extend_from_slice(bytes);
                        } else if !received.is_empty() {
                            output
                                .lock()
                                .unwrap()
                                .push(String::from_utf8_lossy(&received).into_owned());
                        }
                        Ok(())
                    }),
                )
            });
            if result.is_err() {
                let mut error = vybe_runtime::value::Object::new();
                error
                    .properties
                    .insert("__wasi_error".into(), Value::String(Arc::from("io")));
                ctx.resolve_future(id, Value::Object(vybe_runtime::heap::alloc(error)));
            }
            future
        }),
    );
}
