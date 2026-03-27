use super::*;

pub fn register(vm: &mut VM) {
    // System.Net.Http
    let http = ensure_namespace(vm, &["System", "Net", "Http"]);
    set_prop(&http, "get", host_fn_ref(vm, "wasi:http", "get"));
    set_prop(&http, "post", host_fn_ref(vm, "wasi:http", "post"));
    set_prop(&http, "fetch", host_fn_ref(vm, "wasi:http", "fetch"));

    // System.Net.Sockets.TcpClient
    let tcp = ensure_namespace(vm, &["System", "Net", "Sockets", "TcpClient"]);
    set_prop(&tcp, "new", host_fn_ref(vm, "vybe:net", "tcpConnect"));

    // System.Net.Dns
    let dns = ensure_namespace(vm, &["System", "Net", "Dns"]);
    set_prop(&dns, "gethostaddresses", host_fn_ref(vm, "vybe:net", "dnsResolve"));
    set_prop(&dns, "gethostentry", host_fn_ref(vm, "vybe:net", "dnsResolve"));

    // System.IO.StreamReader / StreamWriter
    let sr = ensure_namespace(vm, &["System", "IO", "StreamReader"]);
    set_prop(&sr, "new", host_fn_ref(vm, "vybe:net", "streamReaderNew"));

    let sw = ensure_namespace(vm, &["System", "IO", "StreamWriter"]);
    set_prop(&sw, "new", host_fn_ref(vm, "vybe:net", "streamWriterNew"));

    // System.Security.Cryptography
    let sha = ensure_namespace(vm, &["System", "Security", "Cryptography", "SHA256"]);
    set_prop(&sha, "create", host_fn_ref(vm, "vybe:crypto", "sha256"));

    let md5 = ensure_namespace(vm, &["System", "Security", "Cryptography", "MD5"]);
    set_prop(&md5, "create", host_fn_ref(vm, "vybe:crypto", "md5"));

    // System.Xml.Linq.XDocument
    let xdoc = ensure_namespace(vm, &["System", "Xml", "Linq", "XDocument"]);
    set_prop(&xdoc, "parse", host_fn_ref(vm, "vybe:xml", "parse"));
    set_prop(&xdoc, "load", host_fn_ref(vm, "vybe:xml", "load"));
}
