//! `node:dgram` — Node.js UDP socket module.
//!
//! Reference: <https://nodejs.org/api/dgram.html>.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::VM;

fn make_socket(sock_type: &str) -> Value {
    let mut o = Object::new();
    o.properties.insert("type".into(), Value::String(Arc::from(sock_type)));
    o.properties.insert("fd".into(), Value::I32(-1));
    o.properties.insert("_bindState".into(), Value::I32(0));
    o.properties.insert("readyState".into(), Value::String(Arc::from("open")));
    // I/O methods
    for m in ["bind","close","send","address","connect","disconnect","remoteAddress",
              "setBroadcast","setTTL","setMulticastTTL","setMulticastLoopback",
              "setMulticastInterface","addMembership","dropMembership",
              "addSourceSpecificMembership","dropSourceSpecificMembership",
              "ref","unref","getSendBufferSize","getRecvBufferSize",
              "setSendBufferSize","setRecvBufferSize"] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    // EventEmitter methods
    for m in ["on","once","off","emit","addListener","removeListener","removeAllListeners",
              "listeners","rawListeners","listenerCount","eventNames"] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:dgram", "createSocket", Box::new(|_ctx, args| {
        let sock_type = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(Value::Object(opts)) => {
                let o = opts.lock().unwrap();
                match o.properties.get("type") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => "udp4".to_string(),
                }
            }
            _ => "udp4".to_string(),
        };
        make_socket(&sock_type)
    }));

    vm.register_host_fn("node:dgram", "Socket", Box::new(|_ctx, args| {
        let sock_type = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => "udp4".to_string(),
        };
        make_socket(&sock_type)
    }));
}
