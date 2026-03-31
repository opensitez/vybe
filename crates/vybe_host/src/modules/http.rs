use std::rc::Rc;
use std::cell::RefCell;
use std::io::{Read, Write};
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // Simple HTTP GET — returns response body as string
    vm.register_host_fn("wasi:http", "get", Box::new(|args: &[Value]| {
        let url = s(args, 0);
        match http_request("GET", &url, None) {
            Ok(body) => Value::String(Rc::from(body.as_str())),
            Err(e) => Value::String(Rc::from(format!("Error: {}", e).as_str())),
        }
    }));

    // HTTP POST — body as second arg, returns response
    vm.register_host_fn("wasi:http", "post", Box::new(|args: &[Value]| {
        let url = s(args, 0);
        let body = s(args, 1);
        match http_request("POST", &url, Some(&body)) {
            Ok(resp) => Value::String(Rc::from(resp.as_str())),
            Err(e) => Value::String(Rc::from(format!("Error: {}", e).as_str())),
        }
    }));

    // Full fetch — returns object { status, body, ok }
    vm.register_host_fn("wasi:http", "fetch", Box::new(|args: &[Value]| {
        let url = s(args, 0);
        let method = if args.len() > 1 { s(args, 1) } else { "GET".into() };
        let body = if args.len() > 2 { Some(s(args, 2)) } else { None };

        match http_request(&method, &url, body.as_deref()) {
            Ok(resp_body) => {
                let mut obj = Object::new();
                obj.properties.insert("status".into(), Value::F64(200.0));
                obj.properties.insert("body".into(), Value::String(Rc::from(resp_body.as_str())));
                obj.properties.insert("ok".into(), Value::Bool(true));
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(e) => {
                let mut obj = Object::new();
                obj.properties.insert("status".into(), Value::F64(0.0));
                obj.properties.insert("body".into(), Value::String(Rc::from(format!("{}", e).as_str())));
                obj.properties.insert("ok".into(), Value::Bool(false));
                Value::Object(Rc::new(RefCell::new(obj)))
            }
        }
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}

/// Minimal HTTP client using std::net::TcpStream (no dependencies).
fn http_request(method: &str, url: &str, body: Option<&str>) -> Result<String, String> {
    // Parse URL
    let url = if url.starts_with("http://") { &url[7..] } else if url.starts_with("https://") {
        return Err("HTTPS not supported (use http://)".into());
    } else { url };

    let (host_port, path) = match url.find('/') {
        Some(i) => (&url[..i], &url[i..]),
        None => (url, "/"),
    };

    let (host, port) = match host_port.find(':') {
        Some(i) => (&host_port[..i], host_port[i+1..].parse::<u16>().unwrap_or(80)),
        None => (host_port, 80u16),
    };

    let addr = format!("{}:{}", host, port);
    let mut stream = std::net::TcpStream::connect(&addr)
        .map_err(|e| format!("Connection failed: {}", e))?;

    stream.set_read_timeout(Some(std::time::Duration::from_secs(10)))
        .map_err(|e| format!("Timeout config failed: {}", e))?;

    let content_length = body.map(|b| b.len()).unwrap_or(0);
    let mut request = format!(
        "{} {} HTTP/1.1\r\nHost: {}\r\nConnection: close\r\nContent-Length: {}\r\n\r\n",
        method, path, host, content_length
    );
    if let Some(b) = body {
        request.push_str(b);
    }

    stream.write_all(request.as_bytes())
        .map_err(|e| format!("Write failed: {}", e))?;

    let mut response = String::new();
    let _ = stream.read_to_string(&mut response);

    // Strip HTTP headers — find \r\n\r\n
    match response.find("\r\n\r\n") {
        Some(i) => Ok(response[i+4..].to_string()),
        None => Ok(response),
    }
}
