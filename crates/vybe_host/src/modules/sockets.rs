//! System.Net.Sockets — TcpClient, TcpListener, UdpClient
//! Real socket implementations for networking tests.

use std::collections::HashMap;
use std::io::{Read, Write};
use std::net::{IpAddr, Ipv4Addr, Ipv6Addr, Shutdown, SocketAddr, TcpListener, TcpStream, UdpSocket};
use std::thread;
use std::time::{Duration, Instant};
use std::sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

static NEXT_ID: AtomicU64 = AtomicU64::new(1);

struct SocketState {
    tcp_streams: HashMap<u64, TcpStream>,
    tcp_listeners: HashMap<u64, TcpListener>,
    pending_accepts: HashMap<u64, Vec<TcpStream>>,
    udp_sockets: HashMap<u64, UdpSocket>,
}

fn get_state() -> Arc<Mutex<SocketState>> {
    use std::sync::OnceLock;
    static STATE: OnceLock<Arc<Mutex<SocketState>>> = OnceLock::new();
    STATE.get_or_init(|| Arc::new(Mutex::new(SocketState {
        tcp_streams: HashMap::new(),
        tcp_listeners: HashMap::new(),
        pending_accepts: HashMap::new(),
        udp_sockets: HashMap::new(),
    }))).clone()
}

fn make_obj(type_name: &str, id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from(type_name)));
    obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
    obj.properties.insert("connected".into(), Value::Bool(true));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn get_id(args: &[Value]) -> u64 {
    match args.first() {
        Some(Value::Object(obj)) => {
            obj.lock().unwrap().properties.get("__socket_id").map(|v| v.as_f64() as u64).unwrap_or(0)
        }
        Some(Value::F64(n)) => *n as u64,
        _ => 0,
    }
}

fn tcp_listener_new(args: &[Value]) -> Value {
    let requested_port = args.get(1).or(args.first()).map(|v| v.as_f64() as u16).unwrap_or(8080);
    let addr = format!("127.0.0.1:{}", requested_port);
    match TcpListener::bind(&addr) {
        Ok(listener) => {
            let _ = listener.set_nonblocking(true);
            let bound_port = listener.local_addr().map(|addr| addr.port()).unwrap_or(requested_port);
            let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Arc::from("TcpListener")));
            obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
            obj.properties.insert("port".into(), Value::F64(bound_port as f64));
            get_state().lock().unwrap().tcp_listeners.insert(id, listener);
            Value::Object(Arc::new(Mutex::new(obj)))
        }
        Err(e) => {
            eprintln!("TcpListener bind error on port {}: {}", requested_port, e);
            Value::Null
        }
    }
}

fn tcp_listener_start(_args: &[Value]) -> Value {
    Value::Null
}

fn tcp_listener_stop(args: &[Value]) -> Value {
    let id = get_id(args);
    let state = get_state();
    let mut state = state.lock().unwrap();
    state.tcp_listeners.remove(&id);
    state.pending_accepts.remove(&id);
    Value::Null
}

fn tcp_listener_pending(args: &[Value]) -> Value {
    let id = get_id(args);
    let state = get_state();
    let mut guard = state.lock().unwrap();

    if guard.pending_accepts.get(&id).map(|pending| !pending.is_empty()).unwrap_or(false) {
        return Value::Bool(true);
    }

    let accepted = if let Some(listener) = guard.tcp_listeners.get(&id) {
        match listener.accept() {
            Ok((stream, _addr)) => Some(stream),
            Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => None,
            Err(err) => {
                eprintln!("Pending error: {}", err);
                None
            }
        }
    } else {
        None
    };

    if let Some(stream) = accepted {
        let _ = stream.set_nonblocking(false);
        guard.pending_accepts.entry(id).or_default().push(stream);
        return Value::Bool(true);
    }

    Value::Bool(false)
}

fn tcp_listener_accept(args: &[Value]) -> Value {
    let id = get_id(args);
    let state = get_state();
    let mut guard = state.lock().unwrap();

    let pending_stream = guard.pending_accepts.get_mut(&id).and_then(|pending| {
        if pending.is_empty() {
            None
        } else {
            Some(pending.remove(0))
        }
    });

    let accepted = if let Some(stream) = pending_stream {
        Some(stream)
    } else if let Some(listener) = guard.tcp_listeners.get(&id) {
        match listener.accept() {
            Ok((stream, _addr)) => Some(stream),
            Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => None,
            Err(err) => {
                eprintln!("Accept error: {}", err);
                None
            }
        }
    } else {
        None
    };

        if let Some(stream) = accepted {
            let _ = stream.set_nonblocking(false);
        let client_id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
        guard.tcp_streams.insert(client_id, stream);
        return make_obj("TcpClient", client_id);
    }

    Value::Null
}

fn tcp_connect(args: &[Value]) -> Value {
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
            obj.properties.insert("__type".into(), Value::String(Arc::from("TcpClient")));
            obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
            obj.properties.insert("connected".into(), Value::Bool(true));
            obj.properties.insert("host".into(), Value::String(Arc::from(host.as_str())));
            obj.properties.insert("port".into(), Value::F64(port as f64));
            get_state().lock().unwrap().tcp_streams.insert(id, stream);
            Value::Object(Arc::new(Mutex::new(obj)))
        }
        Err(e) => {
            eprintln!("TcpClient connect error: {}", e);
            Value::Null
        }
    }
}

fn tcp_get_stream(args: &[Value]) -> Value {
    args.first().cloned().unwrap_or(Value::Null)
}

fn tcp_close(args: &[Value]) -> Value {
    let id = get_id(args);
    get_state().lock().unwrap().tcp_streams.remove(&id);
    Value::Null
}

fn udp_new(args: &[Value]) -> Value {
    let port = args.get(1).or(args.first()).map(|v| v.as_f64() as u16).unwrap_or(0);
    let addr = format!("127.0.0.1:{}", port);
    match UdpSocket::bind(&addr) {
        Ok(socket) => {
            let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
            let mut obj = Object::new();
            obj.properties.insert("__type".into(), Value::String(Arc::from("UdpClient")));
            obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
            obj.properties.insert("port".into(), Value::F64(port as f64));
            get_state().lock().unwrap().udp_sockets.insert(id, socket);
            Value::Object(Arc::new(Mutex::new(obj)))
        }
        Err(e) => {
            eprintln!("UdpClient bind error: {}", e);
            Value::Null
        }
    }
}

fn udp_send(args: &[Value]) -> Value {
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
}

fn udp_receive(args: &[Value]) -> Value {
    let id = get_id(args);
    let state = get_state();
    let guard = state.lock().unwrap();
    if let Some(socket) = guard.udp_sockets.get(&id) {
        let mut buf = [0u8; 4096];
        match socket.recv_from(&mut buf) {
            Ok((n, _addr)) => {
                return Value::String(Arc::from(String::from_utf8_lossy(&buf[..n]).as_ref()));
            }
            Err(e) => eprintln!("UDP receive error: {}", e),
        }
    }
    Value::Null
}

fn udp_close(args: &[Value]) -> Value {
    let id = get_id(args);
    get_state().lock().unwrap().udp_sockets.remove(&id);
    Value::Null
}

pub fn register_dotnet_io(vm: &mut VM) {
    vm.register_host_fn("dotnet:io", "streamWriterNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("StreamWriter")));

        match &inner {
            Value::Object(inner_obj) => {
                let inner_locked = inner_obj.lock().unwrap();
                if let Some(Value::String(kind)) = inner_locked.properties.get("__type") {
                    match kind.as_ref() {
                        "OutputStream" => {
                            obj.properties.insert("__output_stream".into(), inner.clone());
                        }
                        "TcpClient" => {
                            if let Some(socket_id) = inner_locked.properties.get("__socket_id") {
                                obj.properties.insert("__output_stream".into(), make_output_stream(socket_id.as_i64() as u64));
                            }
                        }
                        _ => {}
                    }
                }
                if let Some(path) = inner_locked.properties.get("__path") {
                    obj.properties.insert("__path".into(), path.clone());
                }
            }
            Value::String(path) => {
                obj.properties.insert("__path".into(), Value::String(path.clone()));
                obj.properties.insert("__buffer".into(), Value::String(Arc::from("")));
            }
            _ => {}
        }

        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("dotnet:io", "streamWriterWriteLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let text = method_arg(args, 0).map(|v| format!("{}", v)).unwrap_or_default();
        let write_text = format!("{}\n", text);
        write_dotnet_stream_text(args.first(), &write_text)
    }));

    vm.register_host_fn("dotnet:io", "streamWriterWrite", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let text = method_arg(args, 0).map(|v| format!("{}", v)).unwrap_or_default();
        write_dotnet_stream_text(args.first(), &text)
    }));

    vm.register_host_fn("dotnet:io", "streamWriterFlush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        flush_dotnet_stream(args.first())
    }));

    vm.register_host_fn("dotnet:io", "streamWriterClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        close_dotnet_stream(args.first())
    }));

    vm.register_host_fn("dotnet:io", "streamReaderNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("StreamReader")));

        match &inner {
            Value::Object(inner_obj) => {
                let inner_locked = inner_obj.lock().unwrap();
                if let Some(Value::String(kind)) = inner_locked.properties.get("__type") {
                    match kind.as_ref() {
                        "InputStream" => {
                            obj.properties.insert("__input_stream".into(), inner.clone());
                        }
                        "TcpClient" => {
                            if let Some(socket_id) = inner_locked.properties.get("__socket_id") {
                                obj.properties.insert("__input_stream".into(), make_input_stream(socket_id.as_i64() as u64));
                            }
                        }
                        _ => {}
                    }
                }
            }
            Value::String(path) => {
                load_file_reader_state(&mut obj, path);
            }
            _ => {}
        }

        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    vm.register_host_fn("dotnet:io", "streamReaderReadLine", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        read_dotnet_stream_line(args.first())
    }));

    vm.register_host_fn("dotnet:io", "streamReaderReadToEnd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        read_dotnet_stream_to_end(args.first())
    }));

    vm.register_host_fn("dotnet:io", "streamReaderAtEnd", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(Value::Object(obj)) = args.first() else { return Value::Bool(true); };
        let locked = obj.lock().unwrap();
        let pos = locked.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
        if let Some(Value::Object(lines_obj)) = locked.properties.get("__lines") {
            let lines = lines_obj.lock().unwrap();
            if let vybe_bytecode::value::ObjectKind::Array(elems) = &lines.kind {
                return Value::Bool(pos >= elems.len());
            }
        }
        Value::Bool(true)
    }));
}

pub fn register_dotnet_net(vm: &mut VM) {
    vm.register_host_fn("dotnet:net", "dnsGetHostAddresses", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let host = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        value_array(resolve_name_strings(&host).into_iter().map(|address| Value::String(Arc::from(address.as_str()))).collect())
    }));

    vm.register_host_fn("dotnet:net", "dnsGetHostEntry", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let host = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        make_dns_host_entry(&host)
    }));
}

pub fn register_dotnet_sockets(vm: &mut VM) {
    vm.register_host_fn("dotnet:sockets", "tcpListenerNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_new(args)));
    vm.register_host_fn("dotnet:sockets", "tcpListenerStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_start(args)));
    vm.register_host_fn("dotnet:sockets", "tcpListenerStop", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_stop(args)));
    vm.register_host_fn("dotnet:sockets", "tcpListenerPending", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_pending(args)));
    vm.register_host_fn("dotnet:sockets", "tcpListenerAccept", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_accept(args)));
    vm.register_host_fn("dotnet:sockets", "tcpClientNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_connect(args)));
    vm.register_host_fn("dotnet:sockets", "tcpClientGetStream", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_get_stream(args)));
    vm.register_host_fn("dotnet:sockets", "tcpClientClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_close(args)));
    vm.register_host_fn("dotnet:sockets", "udpClientNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_new(args)));
    vm.register_host_fn("dotnet:sockets", "udpClientSend", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_send(args)));
    vm.register_host_fn("dotnet:sockets", "udpClientReceive", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_receive(args)));
    vm.register_host_fn("dotnet:sockets", "udpClientClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_close(args)));
}

pub fn register(vm: &mut VM) {
    // New TcpListener(port)
    vm.register_host_fn("vybe:net", "tcpListenerNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_new(args)));

    // TcpListener.Start() — already bound, no-op
    vm.register_host_fn("vybe:net", "tcpListenerStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_start(args)));

    // TcpListener.Stop()
    vm.register_host_fn("vybe:net", "tcpListenerStop", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_stop(args)));

    // TcpListener.Pending() → check for a queued or newly accepted client without blocking
    vm.register_host_fn("vybe:net", "tcpListenerPending", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_pending(args)));

    // TcpListener.AcceptTcpClient() → TcpClient object
    vm.register_host_fn("vybe:net", "tcpListenerAccept", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_listener_accept(args)));

    // New TcpClient(host, port)
    vm.register_host_fn("vybe:net", "tcpConnect", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_connect(args)));

    // TcpClient.GetStream() → returns self (the stream IS the client in our model)
    vm.register_host_fn("vybe:net", "tcpGetStream", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_get_stream(args)));

    // TcpClient.Close()
    vm.register_host_fn("vybe:net", "tcpClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| tcp_close(args)));

    // Legacy StreamWriter wrapping a TCP stream
    // New StreamWriter(stream) — stream is a TcpClient object
    vm.register_host_fn("vybe:net", "streamWriterNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // args[0] = null (this), args[1] = stream/tcpclient object OR file path
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("StreamWriter")));
        // If inner is a TcpClient, copy its socket id
        if let Value::Object(ref inner_obj) = inner {
            let io = inner_obj.lock().unwrap();
            if let Some(sid) = io.properties.get("__socket_id") {
                obj.properties.insert("__socket_id".into(), sid.clone());
            }
            if let Some(path) = io.properties.get("__path") {
                obj.properties.insert("__path".into(), path.clone());
            }
        } else if let Value::String(ref s) = inner {
            // File path
            obj.properties.insert("__path".into(), Value::String(s.clone()));
            obj.properties.insert("__buffer".into(), Value::String(Arc::from("")));
        }
        Value::Object(Arc::new(Mutex::new(obj)))
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
            let mut o = obj.lock().unwrap();
            let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
            o.properties.insert("__buffer".into(), Value::String(Arc::from(format!("{}{}\n", buf, text).as_str())));
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
            let o = obj.lock().unwrap();
            if let Some(Value::String(path)) = o.properties.get("__path") {
                let buf = o.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
                let _ = std::fs::write(path.as_ref(), &buf);
            }
        }
        Value::Null
    }));

    // Legacy StreamReader wrapping a TCP stream
    vm.register_host_fn("vybe:net", "streamReaderNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let inner = args.get(1).or(args.first()).cloned().unwrap_or(Value::Null);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("StreamReader")));
        if let Value::Object(ref inner_obj) = inner {
            let io = inner_obj.lock().unwrap();
            if let Some(sid) = io.properties.get("__socket_id") {
                obj.properties.insert("__socket_id".into(), sid.clone());
            }
        } else if let Value::String(ref s) = inner {
            // File path
            match std::fs::read_to_string(s.as_ref()) {
                Ok(content) => {
                    let lines: Vec<Value> = content.lines().map(|l| Value::String(Arc::from(l))).collect();
                    obj.properties.insert("__lines".into(), Value::Object(Arc::new(Mutex::new(
                        vybe_bytecode::value::Object::new_array(lines)
                    ))));
                    obj.properties.insert("__pos".into(), Value::F64(0.0));
                }
                Err(_) => {}
            }
        }
        Value::Object(Arc::new(Mutex::new(obj)))
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
                    return Value::String(Arc::from(String::from_utf8_lossy(&line).as_ref()));
                }
                return Value::Null;
            }
        }
        // File-based reader
        if let Some(Value::Object(obj)) = args.first() {
            let mut o = obj.lock().unwrap();
            let pos = o.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
            if let Some(Value::Object(lines_obj)) = o.properties.get("__lines") {
                let lo = lines_obj.lock().unwrap();
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
    vm.register_host_fn("vybe:net", "udpNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_new(args)));

    // UdpClient.Send(data, length, host, port)
    vm.register_host_fn("vybe:net", "udpSend", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_send(args)));

    // UdpClient.Receive() → byte array (as string for now)
    vm.register_host_fn("vybe:net", "udpReceive", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_receive(args)));

    // UdpClient.Close()
    vm.register_host_fn("vybe:net", "udpClose", Box::new(|_ctx: &mut HostContext, args: &[Value]| udp_close(args)));

    // DNS resolve
    vm.register_host_fn("vybe:net", "dnsResolve", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let host = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        value_array(resolve_name_strings(&host).into_iter().map(|address| Value::String(Arc::from(address.as_str()))).collect())
    }));

    register_wasi_io(vm);
    register_wasi_sockets(vm);
}

fn register_wasi_io(vm: &mut VM) {
    vm.register_host_fn("wasi:io/streams", "read", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        match stream_kind(&stream).as_deref() {
            Some("InputStream") => read_stream_bytes(&stream, len, false),
            _ => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-read", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        match stream_kind(&stream).as_deref() {
            Some("InputStream") => read_stream_bytes(&stream, len, true),
            _ => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:io/streams", "skip", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        match read_stream_bytes(&stream, len, false) {
            Value::Object(bytes) => Value::I64(array_len(&bytes) as i64),
            other => other,
        }
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-skip", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        match read_stream_bytes(&stream, len, true) {
            Value::Object(bytes) => Value::I64(array_len(&bytes) as i64),
            other => other,
        }
    }));

    vm.register_host_fn("wasi:io/streams", "subscribe", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        make_pollable(stream)
    }));

    vm.register_host_fn("wasi:io/streams", "check-write", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        match stream_kind(&stream).as_deref() {
            Some("OutputStream") if stream_socket_id(&stream).is_some() => Value::I64(65536),
            _ => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:io/streams", "write", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let contents = method_arg(args, 0).map(bytes_from_value).unwrap_or_default();
        write_stream_bytes(&stream, &contents, false)
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-write-and-flush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let contents = method_arg(args, 0).map(bytes_from_value).unwrap_or_default();
        write_stream_bytes(&stream, &contents, true)
    }));

    vm.register_host_fn("wasi:io/streams", "flush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        flush_stream(&stream)
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-flush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        flush_stream(&stream)
    }));

    vm.register_host_fn("wasi:io/streams", "write-zeroes", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        write_stream_bytes(&stream, &vec![0; len], false)
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-write-zeroes-and-flush", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(stream) = stream_arg(args) else { return Value::Null; };
        let len = method_arg(args, 0).map(value_len_arg).unwrap_or(0);
        write_stream_bytes(&stream, &vec![0; len], true)
    }));

    vm.register_host_fn("wasi:io/streams", "splice", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(dst) = stream_arg(args) else { return Value::Null; };
        let Some(src) = method_arg(args, 0).and_then(as_object_value) else { return Value::Null; };
        let len = method_arg(args, 1).map(value_len_arg).unwrap_or(0);
        splice_streams(&src, &dst, len, false)
    }));

    vm.register_host_fn("wasi:io/streams", "blocking-splice", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(dst) = stream_arg(args) else { return Value::Null; };
        let Some(src) = method_arg(args, 0).and_then(as_object_value) else { return Value::Null; };
        let len = method_arg(args, 1).map(value_len_arg).unwrap_or(0);
        splice_streams(&src, &dst, len, true)
    }));

    vm.register_host_fn("wasi:io/poll", "ready", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(pollable) = pollable_arg(args) else { return Value::Bool(false); };
        Value::Bool(pollable_ready(&pollable))
    }));

    vm.register_host_fn("wasi:io/poll", "block", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(pollable) = pollable_arg(args) else { return Value::Null; };
        block_until_ready(&pollable);
        Value::Null
    }));

    vm.register_host_fn("wasi:io/poll", "poll", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(list) = args.first() else { return value_array(vec![]); };
        let Some(pollables) = value_array_elements(list) else { return value_array(vec![]); };

        let mut ready = collect_ready_indices(&pollables);
        if ready.is_empty() {
            let start = Instant::now();
            while ready.is_empty() && start.elapsed() < Duration::from_secs(1) {
                thread::sleep(Duration::from_millis(1));
                ready = collect_ready_indices(&pollables);
            }
        }

        value_array(ready.into_iter().map(|index| Value::I32(index as i32)).collect())
    }));
}

fn register_wasi_sockets(vm: &mut VM) {
    vm.register_host_fn("wasi:sockets/instance-network", "instance-network", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        make_network_handle()
    }));

    vm.register_host_fn("wasi:sockets/ip-name-lookup", "resolve-addresses", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.last().map(|v| format!("{}", v)).unwrap_or_default();
        let addresses = resolve_name_socket_addrs(&name).into_iter().map(socket_addr_to_value).collect();
        make_resolve_address_stream(addresses)
    }));

    vm.register_host_fn("wasi:sockets/tcp-create-socket", "create-tcp-socket", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let family = args.first().map(address_family_from_value).unwrap_or_else(|| "ipv4".into());
        make_tcp_socket_resource(&family)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "start-bind", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let Some(local_addr_val) = method_arg(args, 1) else { return Value::Null; };
        let Some((host, port, family)) = parse_ip_socket_address(local_addr_val) else { return Value::Null; };
        let bind_addr = format!("{}:{}", host, port);
        match TcpListener::bind(&bind_addr) {
            Ok(listener) => {
                let _ = listener.set_nonblocking(true);
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                get_state().lock().unwrap().tcp_listeners.insert(id, listener);
                let mut obj = socket.lock().unwrap();
                obj.properties.insert("__listener_id".into(), Value::F64(id as f64));
                obj.properties.insert("__state".into(), Value::String(Arc::from("bound")));
                obj.properties.insert("__address_family".into(), Value::String(Arc::from(family.as_str())));
                obj.properties.insert("__local_address".into(), make_ip_socket_address(&host, port, &family));
                Value::Bool(true)
            }
            Err(err) => {
                eprintln!("wasi:sockets tcp.start-bind error: {}", err);
                Value::Null
            }
        }
    }));

    vm.register_host_fn("wasi:sockets/tcp", "finish-bind", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "start-listen", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let mut obj = socket.lock().unwrap();
        obj.properties.insert("__state".into(), Value::String(Arc::from("listening")));
        obj.properties.insert("islistening".into(), Value::Bool(true));
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "finish-listen", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "start-connect", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let Some(remote_addr_val) = method_arg(args, 1) else { return Value::Null; };
        let Some((host, port, family)) = parse_ip_socket_address(remote_addr_val) else { return Value::Null; };
        let remote = format!("{}:{}", host, port);
        match TcpStream::connect(&remote) {
            Ok(stream) => {
                let local_addr = stream.local_addr().ok();
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                get_state().lock().unwrap().tcp_streams.insert(id, stream);
                let mut obj = socket.lock().unwrap();
                obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
                obj.properties.insert("__state".into(), Value::String(Arc::from("connected")));
                obj.properties.insert("__address_family".into(), Value::String(Arc::from(family.as_str())));
                obj.properties.insert("__remote_address".into(), make_ip_socket_address(&host, port, &family));
                if let Some(local_addr) = local_addr {
                    obj.properties.insert("__local_address".into(), socket_addr_to_value(local_addr));
                }
                Value::Bool(true)
            }
            Err(err) => {
                eprintln!("wasi:sockets tcp.start-connect error: {}", err);
                Value::Null
            }
        }
    }));

    vm.register_host_fn("wasi:sockets/tcp", "finish-connect", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let stream_id = socket.lock().unwrap().properties.get("__socket_id").cloned();
        match stream_id {
            Some(Value::F64(id)) => value_array(vec![make_input_stream(id as u64), make_output_stream(id as u64)]),
            _ => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:sockets/tcp", "accept", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let listener_id = socket.lock().unwrap().properties.get("__listener_id").cloned();
        let Some(Value::F64(id)) = listener_id else { return Value::Null; };
        let listener_id = id as u64;
        let state = get_state();
        let mut guard = state.lock().unwrap();

        let pending_stream = guard.pending_accepts.get_mut(&listener_id).and_then(|pending| {
            if pending.is_empty() { None } else { Some(pending.remove(0)) }
        });

        let accepted = if let Some(stream) = pending_stream {
            Some(stream)
        } else if let Some(listener) = guard.tcp_listeners.get(&listener_id) {
            match listener.accept() {
                Ok((stream, _)) => Some(stream),
                Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => None,
                Err(err) => {
                    eprintln!("wasi:sockets tcp.accept error: {}", err);
                    None
                }
            }
        } else {
            None
        };

        if let Some(stream) = accepted {
            let _ = stream.set_nonblocking(false);
            let remote_addr = stream.peer_addr().ok();
            let local_addr = stream.local_addr().ok();
            let client_id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
            guard.tcp_streams.insert(client_id, stream);
            let client = make_tcp_socket_resource("ipv4");
            if let Value::Object(client_obj) = &client {
                let mut obj = client_obj.lock().unwrap();
                obj.properties.insert("__socket_id".into(), Value::F64(client_id as f64));
                obj.properties.insert("__state".into(), Value::String(Arc::from("connected")));
                if let Some(addr) = local_addr {
                    obj.properties.insert("__local_address".into(), socket_addr_to_value(addr));
                }
                if let Some(addr) = remote_addr {
                    obj.properties.insert("__remote_address".into(), socket_addr_to_value(addr));
                }
            }
            return value_array(vec![client, make_input_stream(client_id), make_output_stream(client_id)]);
        }

        Value::Null
    }));

    vm.register_host_fn("wasi:sockets/tcp", "local-address", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__local_address")
    }));

    vm.register_host_fn("wasi:sockets/tcp", "remote-address", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__remote_address")
    }));

    vm.register_host_fn("wasi:sockets/tcp", "is-listening", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match socket_property(args, "islistening") {
            Value::Bool(value) => Value::Bool(value),
            _ => Value::Bool(false),
        }
    }));

    vm.register_host_fn("wasi:sockets/tcp", "address-family", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__address_family")
    }));

    vm.register_host_fn("wasi:sockets/tcp", "set-listen-backlog-size", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let backlog = method_arg(args, 0).cloned().unwrap_or(Value::Null);
        socket.lock().unwrap().properties.insert("__listen_backlog".into(), backlog);
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "subscribe", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("wasi:sockets/tcp", "shutdown", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let shutdown_value = method_arg(args, 0).map(|v| format!("{}", v)).unwrap_or_else(|| "both".into());
        let socket_id = socket.lock().unwrap().properties.get("__socket_id").cloned();
        let Some(Value::F64(id)) = socket_id else { return Value::Null; };
        let state = get_state();
        let mut guard = state.lock().unwrap();
        if let Some(stream) = guard.tcp_streams.get_mut(&(id as u64)) {
            let how = match shutdown_value.as_str() {
                "receive" => Shutdown::Read,
                "send" => Shutdown::Write,
                _ => Shutdown::Both,
            };
            let _ = stream.shutdown(how);
        }
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/udp-create-socket", "create-udp-socket", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let family = args.first().map(address_family_from_value).unwrap_or_else(|| "ipv4".into());
        make_udp_socket_resource(&family)
    }));

    vm.register_host_fn("wasi:sockets/udp", "start-bind", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let Some(local_addr_val) = method_arg(args, 1) else { return Value::Null; };
        let Some((host, port, family)) = parse_ip_socket_address(local_addr_val) else { return Value::Null; };
        let bind_addr = format!("{}:{}", host, port);
        match UdpSocket::bind(&bind_addr) {
            Ok(udp) => {
                let id = NEXT_ID.fetch_add(1, Ordering::Relaxed);
                get_state().lock().unwrap().udp_sockets.insert(id, udp);
                let mut obj = socket.lock().unwrap();
                obj.properties.insert("__socket_id".into(), Value::F64(id as f64));
                obj.properties.insert("__state".into(), Value::String(Arc::from("bound")));
                obj.properties.insert("__address_family".into(), Value::String(Arc::from(family.as_str())));
                obj.properties.insert("__local_address".into(), make_ip_socket_address(&host, port, &family));
                Value::Bool(true)
            }
            Err(err) => {
                eprintln!("wasi:sockets udp.start-bind error: {}", err);
                Value::Null
            }
        }
    }));

    vm.register_host_fn("wasi:sockets/udp", "finish-bind", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Bool(true)
    }));

    vm.register_host_fn("wasi:sockets/udp", "stream", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(socket) = socket_arg(args) else { return Value::Null; };
        let remote_address = method_arg(args, 0).cloned().unwrap_or(Value::Null);
        if !matches!(remote_address, Value::Null) {
            socket.lock().unwrap().properties.insert("__remote_address".into(), remote_address);
        }
        let socket_value = Value::Object(socket.clone());
        value_array(vec![make_incoming_datagram_stream(socket_value.clone()), make_outgoing_datagram_stream(socket_value)])
    }));

    vm.register_host_fn("wasi:sockets/udp", "local-address", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__local_address")
    }));

    vm.register_host_fn("wasi:sockets/udp", "remote-address", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__remote_address")
    }));

    vm.register_host_fn("wasi:sockets/udp", "address-family", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        socket_property(args, "__address_family")
    }));

    vm.register_host_fn("wasi:sockets/udp", "subscribe", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

fn make_network_handle() -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("Network")));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_tcp_socket_resource(family: &str) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("TcpSocket")));
    obj.properties.insert("__state".into(), Value::String(Arc::from("unbound")));
    obj.properties.insert("__address_family".into(), Value::String(Arc::from(family)));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_udp_socket_resource(family: &str) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("UdpSocket")));
    obj.properties.insert("__state".into(), Value::String(Arc::from("unbound")));
    obj.properties.insert("__address_family".into(), Value::String(Arc::from(family)));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_input_stream(socket_id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("InputStream")));
    obj.properties.insert("__socket_id".into(), Value::F64(socket_id as f64));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_output_stream(socket_id: u64) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("OutputStream")));
    obj.properties.insert("__socket_id".into(), Value::F64(socket_id as f64));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_resolve_address_stream(addresses: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("ResolveAddressStream")));
    obj.properties.insert("__addresses".into(), value_array_inner(addresses));
    obj.properties.insert("__pos".into(), Value::F64(0.0));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_incoming_datagram_stream(socket: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("IncomingDatagramStream")));
    obj.properties.insert("__socket".into(), socket);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_outgoing_datagram_stream(socket: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("OutgoingDatagramStream")));
    obj.properties.insert("__socket".into(), socket);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_dns_host_entry(host: &str) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("hostname".into(), Value::String(Arc::from(host)));
    obj.properties.insert(
        "addresslist".into(),
        value_array(resolve_name_strings(host).into_iter().map(|address| Value::String(Arc::from(address.as_str()))).collect()),
    );
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn resolve_name_strings(host: &str) -> Vec<String> {
    resolve_name_socket_addrs(host).into_iter().map(|addr| addr.ip().to_string()).collect()
}

fn resolve_name_socket_addrs(host: &str) -> Vec<SocketAddr> {
    match std::net::ToSocketAddrs::to_socket_addrs(&format!("{}:0", host)) {
        Ok(addrs) => addrs.collect(),
        Err(_) => Vec::new(),
    }
}

fn make_pollable(target: Arc<Mutex<Object>>) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("Pollable")));
    obj.properties.insert("__target".into(), Value::Object(target));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn write_dotnet_stream_text(target: Option<&Value>, text: &str) -> Value {
    let Some(Value::Object(obj)) = target else { return Value::Null; };
    let output_stream = {
        let locked = obj.lock().unwrap();
        locked.properties.get("__output_stream").cloned()
    };

    if let Some(Value::Object(stream)) = output_stream {
        return write_stream_bytes(&stream, text.as_bytes(), false);
    }

    let mut locked = obj.lock().unwrap();
    let buf = locked.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
    locked.properties.insert("__buffer".into(), Value::String(Arc::from(format!("{}{}", buf, text).as_str())));
    Value::Null
}

fn flush_dotnet_stream(target: Option<&Value>) -> Value {
    let Some(Value::Object(obj)) = target else { return Value::Null; };
    let output_stream = {
        let locked = obj.lock().unwrap();
        locked.properties.get("__output_stream").cloned()
    };

    if let Some(Value::Object(stream)) = output_stream {
        return flush_stream(&stream);
    }

    Value::Null
}

fn close_dotnet_stream(target: Option<&Value>) -> Value {
    let Some(Value::Object(obj)) = target else { return Value::Null; };

    let output_stream = {
        let locked = obj.lock().unwrap();
        locked.properties.get("__output_stream").cloned()
    };
    if let Some(Value::Object(stream)) = output_stream {
        let _ = flush_stream(&stream);
    }

    let locked = obj.lock().unwrap();
    if let Some(Value::String(path)) = locked.properties.get("__path") {
        let buf = locked.properties.get("__buffer").map(|v| format!("{}", v)).unwrap_or_default();
        let _ = std::fs::write(path.as_ref(), &buf);
    }
    Value::Null
}

fn load_file_reader_state(obj: &mut Object, path: &Arc<str>) {
    match std::fs::read_to_string(path.as_ref()) {
        Ok(content) => {
            let lines: Vec<Value> = content.lines().map(|line| Value::String(Arc::from(line))).collect();
            obj.properties.insert("__lines".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(lines)))));
            obj.properties.insert("__pos".into(), Value::F64(0.0));
        }
        Err(_) => {}
    }
}

fn read_dotnet_stream_line(target: Option<&Value>) -> Value {
    let Some(Value::Object(obj)) = target else { return Value::Null; };
    let input_stream = {
        let locked = obj.lock().unwrap();
        locked.properties.get("__input_stream").cloned()
    };

    if let Some(Value::Object(stream)) = input_stream {
        let mut line = Vec::new();
        loop {
            let chunk = read_stream_bytes(&stream, 1, true);
            let bytes = bytes_from_value(&chunk);
            if bytes.is_empty() {
                break;
            }
            let byte = bytes[0];
            if byte == b'\n' {
                break;
            }
            if byte != b'\r' {
                line.push(byte);
            }
        }
        if line.is_empty() {
            return Value::Null;
        }
        return Value::String(Arc::from(String::from_utf8_lossy(&line).as_ref()));
    }

    let mut locked = obj.lock().unwrap();
    let pos = locked.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
    if let Some(Value::Object(lines_obj)) = locked.properties.get("__lines") {
        let lines = lines_obj.lock().unwrap();
        if let vybe_bytecode::value::ObjectKind::Array(elems) = &lines.kind {
            if pos < elems.len() {
                let line = elems[pos].clone();
                let new_pos = pos + 1;
                let at_end = new_pos >= elems.len();
                drop(lines);
                locked.properties.insert("__pos".into(), Value::F64(new_pos as f64));
                locked.properties.insert("endofstream".into(), Value::Bool(at_end));
                return line;
            }
        }
    }
    locked.properties.insert("endofstream".into(), Value::Bool(true));
    Value::Null
}

fn read_dotnet_stream_to_end(target: Option<&Value>) -> Value {
    let Some(Value::Object(obj)) = target else { return Value::Null; };
    let input_stream = {
        let locked = obj.lock().unwrap();
        locked.properties.get("__input_stream").cloned()
    };

    if let Some(Value::Object(stream)) = input_stream {
        let mut bytes = Vec::new();
        loop {
            let chunk = read_stream_bytes(&stream, 4096, true);
            let mut part = bytes_from_value(&chunk);
            if part.is_empty() {
                break;
            }
            bytes.append(&mut part);
            if bytes.len() >= 4096 && bytes.len() % 4096 != 0 {
                break;
            }
        }
        return Value::String(Arc::from(String::from_utf8_lossy(&bytes).as_ref()));
    }

    let mut locked = obj.lock().unwrap();
    let pos = locked.properties.get("__pos").map(|v| v.as_f64() as usize).unwrap_or(0);
    if let Some(Value::Object(lines_obj)) = locked.properties.get("__lines").cloned() {
        let lines = lines_obj.lock().unwrap();
        if let vybe_bytecode::value::ObjectKind::Array(elems) = &lines.kind {
            let remaining = elems[pos..].iter().map(|value| format!("{}", value)).collect::<Vec<_>>().join("\n");
            let total = elems.len();
            drop(lines);
            locked.properties.insert("__pos".into(), Value::F64(total as f64));
            locked.properties.insert("endofstream".into(), Value::Bool(true));
            return Value::String(Arc::from(remaining.as_str()));
        }
    }
    locked.properties.insert("endofstream".into(), Value::Bool(true));
    Value::Null
}

fn elems_len(lines_obj: &Arc<Mutex<Object>>) -> usize {
    let lines = lines_obj.lock().unwrap();
    if let vybe_bytecode::value::ObjectKind::Array(elems) = &lines.kind {
        elems.len()
    } else {
        0
    }
}

fn socket_arg(args: &[Value]) -> Option<Arc<Mutex<Object>>> {
    match args.first() {
        Some(Value::Object(obj)) => Some(obj.clone()),
        _ => None,
    }
}

fn method_arg<'a>(args: &'a [Value], index: usize) -> Option<&'a Value> {
    let offset = usize::from(matches!(args.first(), Some(Value::Object(_))));
    args.get(offset + index)
}

fn socket_property(args: &[Value], key: &str) -> Value {
    match args.first() {
        Some(Value::Object(obj)) => obj.lock().unwrap().properties.get(key).cloned().unwrap_or(Value::Null),
        _ => Value::Null,
    }
}

fn stream_arg(args: &[Value]) -> Option<Arc<Mutex<Object>>> {
    socket_arg(args)
}

fn pollable_arg(args: &[Value]) -> Option<Arc<Mutex<Object>>> {
    socket_arg(args)
}

fn as_object_value(value: &Value) -> Option<Arc<Mutex<Object>>> {
    match value {
        Value::Object(obj) => Some(obj.clone()),
        _ => None,
    }
}

fn stream_kind(stream: &Arc<Mutex<Object>>) -> Option<String> {
    match stream.lock().unwrap().properties.get("__type") {
        Some(Value::String(kind)) => Some(kind.to_string()),
        _ => None,
    }
}

fn stream_socket_id(stream: &Arc<Mutex<Object>>) -> Option<u64> {
    match stream.lock().unwrap().properties.get("__socket_id") {
        Some(Value::F64(id)) => Some(*id as u64),
        Some(Value::I64(id)) => Some(*id as u64),
        Some(Value::I32(id)) => Some(*id as u64),
        _ => None,
    }
}

fn value_len_arg(value: &Value) -> usize {
    let len = value.as_i64();
    len.clamp(0, 64 * 1024) as usize
}

fn byte_array(bytes: &[u8]) -> Value {
    value_array(bytes.iter().map(|byte| Value::I32(*byte as i32)).collect())
}

fn bytes_from_value(value: &Value) -> Vec<u8> {
    match value {
        Value::String(text) => text.as_bytes().to_vec(),
        Value::Object(array) => {
            let array = array.lock().unwrap();
            if let vybe_bytecode::value::ObjectKind::Array(elems) = &array.kind {
                elems.iter().map(|elem| elem.as_i32().clamp(0, 255) as u8).collect()
            } else {
                Vec::new()
            }
        }
        _ => Vec::new(),
    }
}

fn read_stream_bytes(stream: &Arc<Mutex<Object>>, len: usize, blocking: bool) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else { return Value::Null; };
    if len == 0 {
        return value_array(vec![]);
    }

    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else { return Value::Null; };

    let mut buf = vec![0u8; len];
    if !blocking {
        let _ = tcp_stream.set_nonblocking(true);
    }
    let result = match tcp_stream.read(&mut buf) {
        Ok(count) => byte_array(&buf[..count]),
        Err(err) if !blocking && err.kind() == std::io::ErrorKind::WouldBlock => value_array(vec![]),
        Err(err) => {
            eprintln!("wasi:io/streams read error: {}", err);
            Value::Null
        }
    };
    if !blocking {
        let _ = tcp_stream.set_nonblocking(false);
    }
    result
}

fn write_stream_bytes(stream: &Arc<Mutex<Object>>, contents: &[u8], flush: bool) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else { return Value::Null; };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else { return Value::Null; };

    match tcp_stream.write_all(contents) {
        Ok(()) => {
            if flush {
                let _ = tcp_stream.flush();
            }
            Value::Null
        }
        Err(err) => {
            eprintln!("wasi:io/streams write error: {}", err);
            Value::Null
        }
    }
}

fn flush_stream(stream: &Arc<Mutex<Object>>) -> Value {
    let Some(socket_id) = stream_socket_id(stream) else { return Value::Null; };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else { return Value::Null; };
    match tcp_stream.flush() {
        Ok(()) => Value::Null,
        Err(err) => {
            eprintln!("wasi:io/streams flush error: {}", err);
            Value::Null
        }
    }
}

fn splice_streams(src: &Arc<Mutex<Object>>, dst: &Arc<Mutex<Object>>, len: usize, blocking: bool) -> Value {
    let bytes = read_stream_bytes(src, len, blocking);
    let Value::Object(array) = &bytes else { return Value::Null; };
    let contents = bytes_from_value(&Value::Object(array.clone()));
    if write_stream_bytes(dst, &contents, blocking).type_tag() == "null" {
        Value::I64(contents.len() as i64)
    } else {
        Value::Null
    }
}

fn value_array_elements(value: &Value) -> Option<Vec<Value>> {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let vybe_bytecode::value::ObjectKind::Array(elements) = &obj.kind {
                Some(elements.clone())
            } else {
                None
            }
        }
        _ => None,
    }
}

fn array_len(array: &Arc<Mutex<Object>>) -> usize {
    let array = array.lock().unwrap();
    if let vybe_bytecode::value::ObjectKind::Array(elements) = &array.kind {
        elements.len()
    } else {
        0
    }
}

fn collect_ready_indices(pollables: &[Value]) -> Vec<usize> {
    pollables.iter().enumerate().filter_map(|(index, value)| {
        let pollable = as_object_value(value)?;
        pollable_ready(&pollable).then_some(index)
    }).collect()
}

fn pollable_ready(pollable: &Arc<Mutex<Object>>) -> bool {
    let target = {
        let pollable = pollable.lock().unwrap();
        pollable.properties.get("__target").cloned().unwrap_or(Value::Null)
    };
    let Some(target) = as_object_value(&target) else {
        return stream_or_socket_ready(pollable);
    };
    stream_or_socket_ready(&target)
}

fn block_until_ready(pollable: &Arc<Mutex<Object>>) {
    let start = Instant::now();
    while !pollable_ready(pollable) && start.elapsed() < Duration::from_secs(1) {
        thread::sleep(Duration::from_millis(1));
    }
}

fn stream_or_socket_ready(resource: &Arc<Mutex<Object>>) -> bool {
    match stream_kind(resource).as_deref() {
        Some("InputStream") => input_stream_ready(resource),
        Some("OutputStream") => stream_socket_id(resource).is_some(),
        Some("TcpSocket") => tcp_socket_ready(resource),
        Some("Pollable") => pollable_ready(resource),
        _ => false,
    }
}

fn input_stream_ready(stream: &Arc<Mutex<Object>>) -> bool {
    let Some(socket_id) = stream_socket_id(stream) else { return false; };
    let state = get_state();
    let mut guard = state.lock().unwrap();
    let Some(tcp_stream) = guard.tcp_streams.get_mut(&socket_id) else { return false; };

    let _ = tcp_stream.set_nonblocking(true);
    let mut buf = [0u8; 1];
    let ready = match tcp_stream.peek(&mut buf) {
        Ok(_) => true,
        Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => false,
        Err(_) => true,
    };
    let _ = tcp_stream.set_nonblocking(false);
    ready
}

fn tcp_socket_ready(socket: &Arc<Mutex<Object>>) -> bool {
    let listener_id = {
        let socket = socket.lock().unwrap();
        socket.properties.get("__listener_id").cloned()
    };
    let Some(Value::F64(listener_id)) = listener_id else {
        return stream_socket_id(socket).is_some();
    };
    let listener_id = listener_id as u64;

    let state = get_state();
    let mut guard = state.lock().unwrap();
    if guard.pending_accepts.get(&listener_id).map(|pending| !pending.is_empty()).unwrap_or(false) {
        return true;
    }

    let accepted = if let Some(listener) = guard.tcp_listeners.get(&listener_id) {
        match listener.accept() {
            Ok((stream, _)) => Some(stream),
            Err(err) if err.kind() == std::io::ErrorKind::WouldBlock => None,
            Err(_) => None,
        }
    } else {
        None
    };

    if let Some(stream) = accepted {
        guard.pending_accepts.entry(listener_id).or_default().push(stream);
        return true;
    }

    false
}

fn address_family_from_value(value: &Value) -> String {
    let raw = format!("{}", value).to_lowercase();
    if raw.contains('6') { "ipv6".into() } else { "ipv4".into() }
}

fn parse_ip_socket_address(value: &Value) -> Option<(String, u16, String)> {
    match value {
        Value::String(s) => parse_socket_addr_str(s),
        Value::Object(obj) => {
            let props = &obj.lock().unwrap().properties;
            if let Some(Value::Object(inner)) = props.get("ipv4").or_else(|| props.get("ipv6")) {
                let family = if props.contains_key("ipv6") { "ipv6" } else { "ipv4" };
                return parse_ip_socket_record(&inner.lock().unwrap().properties, family);
            }
            parse_ip_socket_record(props, "ipv4")
        }
        _ => None,
    }
}

fn parse_ip_socket_record(props: &HashMap<String, Value>, default_family: &str) -> Option<(String, u16, String)> {
    let port = props.get("port")?.as_f64() as u16;
    if let Some(Value::String(address)) = props.get("address") {
        return Some((address.to_string(), port, default_family.to_string()));
    }
    if let Some(Value::Object(address)) = props.get("address") {
        let address_obj = address.lock().unwrap();
        if let vybe_bytecode::value::ObjectKind::Array(elems) = &address_obj.kind {
            let family = if elems.len() == 4 { "ipv4" } else { "ipv6" };
            let host = if elems.len() == 4 {
                format!("{}.{}.{}.{}", elems[0].as_i32(), elems[1].as_i32(), elems[2].as_i32(), elems[3].as_i32())
            } else {
                elems.iter().map(|elem| format!("{:x}", elem.as_i32())).collect::<Vec<_>>().join(":")
            };
            return Some((host, port, family.into()));
        }
    }
    None
}

fn parse_socket_addr_str(input: &str) -> Option<(String, u16, String)> {
    if let Ok(addr) = input.parse::<SocketAddr>() {
        let family = match addr.ip() {
            IpAddr::V4(_) => "ipv4",
            IpAddr::V6(_) => "ipv6",
        };
        return Some((addr.ip().to_string(), addr.port(), family.into()));
    }
    if let Some((host, port)) = input.rsplit_once(':') {
        if let Ok(port) = port.parse::<u16>() {
            let family = if host.contains(':') { "ipv6" } else { "ipv4" };
            return Some((host.to_string(), port, family.into()));
        }
    }
    None
}

fn socket_addr_to_value(addr: SocketAddr) -> Value {
    make_ip_socket_address(&addr.ip().to_string(), addr.port(), if addr.is_ipv6() { "ipv6" } else { "ipv4" })
}

fn make_ip_socket_address(host: &str, port: u16, family: &str) -> Value {
    let mut obj = Object::new();
    obj.properties.insert("family".into(), Value::String(Arc::from(family)));
    obj.properties.insert("port".into(), Value::F64(port as f64));
    obj.properties.insert("address".into(), ip_address_value(host, family));
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn ip_address_value(host: &str, family: &str) -> Value {
    let values = if family == "ipv6" {
        host.parse::<Ipv6Addr>()
            .map(|addr| addr.segments().into_iter().map(|segment| Value::I32(segment as i32)).collect())
            .unwrap_or_else(|_| vec![Value::I32(0); 8])
    } else {
        host.parse::<Ipv4Addr>()
            .map(|addr| addr.octets().into_iter().map(|octet| Value::I32(octet as i32)).collect())
            .unwrap_or_else(|_| vec![Value::I32(127), Value::I32(0), Value::I32(0), Value::I32(1)])
    };
    value_array_inner(values)
}

fn value_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn value_array_inner(elements: Vec<Value>) -> Value {
    value_array(elements)
}
