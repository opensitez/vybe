//! System.Net.Sockets — TcpClient, TcpListener, UdpClient
//! Ported from vybe_runtime/src/builtins/networking.rs

use std::cell::RefCell;
use std::rc::Rc;
use std::io::{Read, Write};
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // TcpClient.New(host, port) → connected client object
    vm.register_host_fn("vybe:net", "tcpConnect", Box::new(|args: &[Value]| {
        let host = s(args, 0);
        let port = args.get(1).map(|v| v.as_f64() as u16).unwrap_or(80);
        match std::net::TcpStream::connect(format!("{}:{}", host, port)) {
            Ok(stream) => {
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Rc::from("TcpClient")));
                obj.properties.insert("connected".into(), Value::Bool(true));
                obj.properties.insert("host".into(), Value::String(Rc::from(host.as_str())));
                obj.properties.insert("port".into(), Value::F64(port as f64));
                // Store the stream in a leaked box (simplified — no GC)
                let stream_ptr = Box::into_raw(Box::new(stream)) as u64;
                obj.properties.insert("__stream_ptr".into(), Value::F64(stream_ptr as f64));
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(e) => {
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Rc::from("TcpClient")));
                obj.properties.insert("connected".into(), Value::Bool(false));
                obj.properties.insert("error".into(), Value::String(Rc::from(format!("{}", e).as_str())));
                Value::Object(Rc::new(RefCell::new(obj)))
            }
        }
    }));

    // Send data over TCP
    vm.register_host_fn("vybe:net", "tcpSend", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let ptr = o.properties.get("__stream_ptr").map(|v| v.as_f64() as u64).unwrap_or(0);
            if ptr != 0 {
                let stream = unsafe { &mut *(ptr as *mut std::net::TcpStream) };
                let data = s(args, 1);
                let _ = stream.write_all(data.as_bytes());
                let _ = stream.flush();
            }
        }
        Value::Null
    }));

    // Receive data from TCP
    vm.register_host_fn("vybe:net", "tcpReceive", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let ptr = o.properties.get("__stream_ptr").map(|v| v.as_f64() as u64).unwrap_or(0);
            if ptr != 0 {
                let stream = unsafe { &mut *(ptr as *mut std::net::TcpStream) };
                let size = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(4096);
                let mut buf = vec![0u8; size];
                match stream.read(&mut buf) {
                    Ok(n) => return Value::String(Rc::from(String::from_utf8_lossy(&buf[..n]).as_ref())),
                    Err(_) => return Value::String(Rc::from("")),
                }
            }
        }
        Value::String(Rc::from(""))
    }));

    // Close TCP connection
    vm.register_host_fn("vybe:net", "tcpClose", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            let ptr = o.properties.get("__stream_ptr").map(|v| v.as_f64() as u64).unwrap_or(0);
            if ptr != 0 {
                unsafe { drop(Box::from_raw(ptr as *mut std::net::TcpStream)); }
                o.properties.insert("__stream_ptr".into(), Value::F64(0.0));
                o.properties.insert("connected".into(), Value::Bool(false));
            }
        }
        Value::Null
    }));

    // DNS resolve
    vm.register_host_fn("vybe:net", "dnsResolve", Box::new(|args: &[Value]| {
        let host = s(args, 0);
        match std::net::ToSocketAddrs::to_socket_addrs(&format!("{}:0", host)) {
            Ok(addrs) => {
                let ips: Vec<Value> = addrs
                    .map(|a| Value::String(Rc::from(a.ip().to_string().as_str())))
                    .collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(ips))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        }
    }));

    // StreamReader — read file line by line
    vm.register_host_fn("vybe:net", "streamReaderNew", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read_to_string(&path) {
            Ok(content) => {
                let lines: Vec<Value> = content.lines()
                    .map(|l| Value::String(Rc::from(l)))
                    .collect();
                let mut obj = Object::new_array(lines);
                obj.properties.insert("__type".into(), Value::String(Rc::from("StreamReader")));
                obj.properties.insert("__pos".into(), Value::F64(0.0));
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(_) => Value::Null,
        }
    }));

    // StreamReader.ReadLine()
    vm.register_host_fn("vybe:net", "streamReaderReadLine", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            let pos = o.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
            if let vybe_bytecode::value::ObjectKind::Array(ref elems) = o.kind {
                if pos < elems.len() {
                    let line = elems[pos].clone();
                    drop(elems);
                    o.properties.insert("__pos".into(), Value::F64((pos + 1) as f64));
                    return line;
                }
            }
        }
        Value::Null
    }));

    // StreamWriter — write to file
    vm.register_host_fn("vybe:net", "streamWriterNew", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("StreamWriter")));
        obj.properties.insert("__path".into(), Value::String(Rc::from(path.as_str())));
        obj.properties.insert("__buffer".into(), Value::String(Rc::from("")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    vm.register_host_fn("vybe:net", "streamWriterWriteLine", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let text = s(args, 1);
            let mut o = obj.borrow_mut();
            let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            o.properties.insert("__buffer".into(), Value::String(Rc::from(format!("{}{}\n", buf, text).as_str())));
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:net", "streamWriterClose", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let path = o.properties.get("__path").map(|v| format!("{}", v)).unwrap_or_default();
            let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            let _ = std::fs::write(&path, &buf);
        }
        Value::Null
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
