use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::ObjectKind;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:json", "stringify", Box::new(|args: &[Value]| {
        Value::String(std::rc::Rc::from(stringify(args.first().unwrap_or(&Value::Null)).as_str()))
    }));
}

fn stringify(v: &Value) -> String {
    match v {
        Value::Null => "null".into(),
        Value::Bool(b) => b.to_string(),
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_nan() || n.is_infinite() { "null".into() }
            else if *n == (*n as i64) as f64 { format!("{}", *n as i64) }
            else { format!("{}", n) }
        }
        Value::String(s) => format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\"").replace('\n', "\\n")),
        Value::Object(obj) => {
            let o = obj.borrow();
            match &o.kind {
                ObjectKind::Array(elems) => {
                    let parts: Vec<String> = elems.iter().map(|e| stringify(e)).collect();
                    format!("[{}]", parts.join(","))
                }
                _ => {
                    let parts: Vec<String> = o.properties.iter()
                        .map(|(k, v)| format!("\"{}\":{}", k, stringify(v)))
                        .collect();
                    format!("{{{}}}", parts.join(","))
                }
            }
        }
    }
}
