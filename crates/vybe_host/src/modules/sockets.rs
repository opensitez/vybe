//! System.Net.Sockets — TcpClient, TcpListener, UdpClient
//! Real socket implementations for networking tests.

use std::cell::RefCell;
use std::collections::HashMap;
use std::io::{Read, Write};
use std::net::{TcpStream, TcpListener, UdpSocket};
use std::rc::Rc;
use std::sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

static NEXT_ID: AtomicU64 = AtomicU64::new(1);

struct SocketState {
    tcp_streams: HashMap<u64, TcpStream>,
    tcp_listeners: HashMap<u64, TcpListener>,
    udp_sockets: HashMap<u64, UdpSocket>,
}

fn get_state() -> Arc<Mutex<SocketState>> {
    use std::sync::OnceLock;
    static STATE: OnceLock<Arc<Mutex<SocketState>>> = OnceLock::new();
    STATE.get_or_init(|| Arc::new(Mutex::new(SocketState {
        tcp_streams: HashMap::new(),
        tcp_listeners: HashMap::new(),
        udp_sockets: HashMap::new(),
    }))).clone()
}

fn make_obj(type_name: &str, id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Rc::from(type_name)));
    obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
    obj.properties.insert("connected".into(), Value::Bool(true));
    Value::Object(Rc::new(RefCell::new(obj)))
}

fn get_id(args: &[Value]) -> u64 {
    match args.first() {
        Some(Value::Object(obj)) => {
            obj.borrow().properties.get("__socket_id").map(|v| v.as_f64() as u64).unwrap_or(0)
        }
        Some(Value::F64(n)) => *n as u64,
        _ => 0,
    }
}

pub fn register(vm: &mut VM) {
    // New TcpListener(port)
    vm.register_host_fn("vybe:net", "tcpListenerNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // args[0] = null (this from New), args[1] = port
        let port = args.get(1).or(args.first()).map(|v| v.as_f64() as u16).unwrap_or(8080);
        let addr = format!("127.0.0.1:{}", port);
        match TcpListener::bind(&addr) {
            Ok(listener) => {
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Rc::from("TcpListener")));
                obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
                obj.properties.insert("port".into(), Value::F64(port as f64));
                get_state().lock().unwrap().tcp_listeners.insert(id, listener);
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(e) => {
                eprintln!("TcpListener bind error on port {}: {}", port, e);
                Value::Null
            }
        }
    }));

    // TcpListener.Start() — already bound, no-op
    vm.register_host_fn("vybe:net", "tcpListenerStart", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Null
    }));

    // TcpListener.Stop()
    vm.register_host_fn("vybe:net", "tcpListenerStop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        get_state().lock().unwrap().tcp_listeners.remove(&id);
        Value::Null
    }));

    // TcpListener.AcceptTcpClient() → TcpClient object
    vm.register_host_fn("vybe:net", "tcpListenerAccept", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        if let Some(listener) = guard.tcp_listeners.get(&id) {
            match listener.accept() {
                Ok((stream, _addr)) => {
                    let client_id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                    drop(guard);
                    get_state().lock().unwrap().tcp_streams.insert(client_id, stream);
                    return make_obj("TcpClient", client_id);
                }
                Err(e) => eprintln!("Accept error: {}", e),
            }
        }
        Value::Null
    }));

    // New TcpClient(host, port)
    vm.register_host_fn("vybe:net", "tcpConnect", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // args might be [null, host, port] from New or [host, port] direct
        let (host, port) = if args.len() >= 3 {
            (format!("{}", args[1]), args[2].as_f64() as u16)
        } else if args.len() >= 2 {
            (format!("{}", args[0]), args[1].as_f64() as u16)
        } else {
            ("127.0.0.1".into(), 80)
        };
        let addr = format!("{}:{}", host, port);
        match TcpStream::connect(&addr) {
            Ok(stream) => {
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Rc::from("TcpClient")));
                obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
                obj.properties.insert("connected".into(), Value::Bool(true));
                obj.properties.insert("host".into(), Value::String(Rc::from(host.as_str())));
                obj.properties.insert("port".into(), Value::F64(port as f64));
                get_state().lock().unwrap().tcp_streams.insert(id, stream);
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(e) => {
                eprintln!("TcpClient connect error: {}", e);
                Value::Null
            }
        }
    }));

    // TcpClient.GetStream() → returns self (the stream IS the client in our model)
    vm.register_host_fn("vybe:net", "tcpGetStream", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));

    // TcpClient.Close()
    vm.register_host_fn("vybe:net", "tcpClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        get_state().lock().unwrap().tcp_streams.remove(&id);
        Value::Null
    }));

    // StreamWriter wrapping a TCP stream
    // New StreamWriter(stream) — stream is a TcpClient object
    vm.register_host_fn("vybe:net", "streamWriterNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // args[0] = null (this), args[1] = stream/tcpclient object OR file path
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("StreamWriter")));
        // If inner is a TcpClient, copy its socket id
        if let Value::Object(ref inner_obj) = inner {
            let io = inner_obj.borrow();
            if let Some(sid) = io.properties.get("__socket_id") {
                obj.properties.insert("__socket_id".into(), sid.clone());
            }
            if let Some(path) = io.properties.get("__path") {
                obj.properties.insert("__path".into(), path.clone());
            }
        } else if let Value::String(ref s) = inner {
            // File path
            obj.properties.insert("__path".into(), Value::String(s.clone()));
            obj.properties.insert("__buffer".into(), Value::String(Rc::from("")));
        }
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // StreamWriter.WriteLine(text)
    vm.register_host_fn("vybe:net", "streamWriterWriteLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let text = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let id = get_id(args);
        if id > 0 {
            // TCP stream write
            let state = get_state();
            let mut guard = state.lock().unwrap();
            if let Some(stream) = guard.tcp_streams.get_mut(&id) {
                let _ = write!(stream, "{}\n", text);
                let _ = stream.flush();
            }
        } else if let Some(Value::Object(obj)) = args.first() {
            // File buffer
            let mut o = obj.borrow_mut();
            let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            o.properties.insert("__buffer".into(), Value::String(Rc::from(format!("{}{}\n", buf, text).as_str())));
        }
        Value::Null
    }));

    // StreamWriter.Flush()
    vm.register_host_fn("vybe:net", "streamWriterFlush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        if id > 0 {
            let state = get_state();
            let mut guard = state.lock().unwrap();
            if let Some(stream) = guard.tcp_streams.get_mut(&id) {
                let _ = stream.flush();
            }
        }
        Value::Null
    }));

    // StreamWriter.Close()
    vm.register_host_fn("vybe:net", "streamWriterClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            if let Some(Value::String(path)) = o.properties.get("__path") {
                let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
                let _ = std::fs::write(path.as_ref(), &buf);
            }
        }
        Value::Null
    }));

    // StreamReader wrapping a TCP stream
    vm.register_host_fn("vybe:net", "streamReaderNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("StreamReader")));
        if let Value::Object(ref inner_obj) = inner {
            let io = inner_obj.borrow();
            if let Some(sid) = io.properties.get("__socket_id") {
                obj.properties.insert("__socket_id".into(), sid.clone());
            }
        } else if let Value::String(ref s) = inner {
            // File path
            match std::fs::read_to_string(s.as_ref()) {
                Ok(content) => {
                    let lines: Vec<Value> = content.lines().map(|l| Value::String(Rc::from(l))).collect();
                    obj.properties.insert("__lines".into(), Value::Object(Rc::new(RefCell::new(
                        vybe_bytecode::value::Object::new_array(lines)
                    ))));
                    obj.properties.insert("__pos".into(), Value::F64(0.0));
                }
                Err(_) => {}
            }
        }
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // StreamReader.ReadLine()
    vm.register_host_fn("vybe:net", "streamReaderReadLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        if id > 0 {
            // TCP stream read — read byte by byte until \n
            let state = get_state();
            let mut guard = state.lock().unwrap();
            if let Some(stream) = guard.tcp_streams.get_mut(&id) {
                let mut line = Vec::new();
                let mut buf = [0u8; 1];
                loop {
                    match stream.read(&mut buf) {
                        Ok(0) => break,
                        Ok(1) => {
                            if buf[0] == b'\n' { break; }
                            if buf[0] != b'\r' { line.push(buf[0]); }
                        }
                        _ => break,
                    }
                }
                if !line.is_empty() {
                    return Value::String(Rc::from(String::from_utf8_lossy(&line).as_ref()));
                }
                return Value::Null;
            }
        }
        // File-based reader
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.borrow_mut();
            let pos = o.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
            if let Some(Value::Object(lines_obj)) = o.properties.get("__lines") {
                let lo = lines_obj.borrow();
                if let vybe_bytecode::value::ObjectKind::Array(ref elems) = lo.kind {
                    if pos < elems.len() {
                        let line = elems[pos].clone();
                        drop(lo);
                        o.properties.insert("__pos".into(), Value::F64((pos + 1) as f64));
                        return line;
                    }
                }
            }
        }
        Value::Null
    }));

    // New UdpClient(port)
    vm.register_host_fn("vybe:net", "udpNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let port = args.get(1).or(args.first()).map(|v| v.as_f64() as u16).unwrap_or(0);
        let addr = format!("127.0.0.1:{}", port);
        match UdpSocket::bind(&addr) {
            Ok(socket) => {
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                let mut obj = Object::new();
                obj.properties.insert("__type".into(), Value::String(Rc::from("UdpClient")));
                obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
                obj.properties.insert("port".into(), Value::F64(port as f64));
                get_state().lock().unwrap().udp_sockets.insert(id, socket);
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(e) => {
                eprintln!("UdpClient bind error: {}", e);
                Value::Null
            }
        }
    }));

    // UdpClient.Send(data, length, host, port)
    vm.register_host_fn("vybe:net", "udpSend", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        let data = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let host = args.get(3).map(|v| format!("{}", v)).unwrap_or_else(|| "127.0.0.1".into());
        let port = args.get(4).map(|v| v.as_f64() as u16).unwrap_or(0);
        let state = get_state();
        let guard = state.lock().unwrap();
        if let Some(socket) = guard.udp_sockets.get(&id) {
            let _ = socket.send_to(data.as_bytes(), format!("{}:{}", host, port));
        }
        Value::Null
    }));

    // UdpClient.Receive() → byte array (as string for now)
    vm.register_host_fn("vybe:net", "udpReceive", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        let state = get_state();
        let guard = state.lock().unwrap();
        if let Some(socket) = guard.udp_sockets.get(&id) {
            let mut buf = [0u8; 4096];
            match socket.recv_from(&mut buf) {
                Ok((n, _addr)) => {
                    return Value::String(Rc::from(String::from_utf8_lossy(&buf[..n]).as_ref()));
                }
                Err(e) => eprintln!("UDP receive error: {}", e),
            }
        }
        Value::Null
    }));

    // UdpClient.Close()
    vm.register_host_fn("vybe:net", "udpClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let id = get_id(args);
        get_state().lock().unwrap().udp_sockets.remove(&id);
        Value::Null
    }));

    // DNS resolve
    vm.register_host_fn("vybe:net", "dnsResolve", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let host = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match std::net::ToSocketAddrs::to_socket_addrs(&format!("{}:0", host)) {
            Ok(addrs) => {
                let ips: Vec<Value> = addrs
                    .map(|a| Value::String(Rc::from(a.ip().to_string().as_str())))
                    .collect();
                Value::Object(Rc::new(RefCell::new(vybe_bytecode::value::Object::new_array(ips))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(vybe_bytecode::value::Object::new_array(vec![])))),
        }
    }));
}
