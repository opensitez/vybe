use super::helpers::{compile_ok, parse_ok};



// ── DateTime (same as VB DateTime, Python datetime, PHP DateTime) ──
#[test] fn datetime_now() { compile_ok("var now = DateTime.now();"); }
#[test] fn datetime_new() { compile_ok("var dt = DateTime(2024, 1, 15);"); }

// ── StringBuffer (same as VB/C# StringBuilder, PHP StringBuilder) ──
#[test] fn string_buffer() { compile_ok("var sb = StringBuffer(); sb.write('hello');"); }

// ── Random (same as VB Random, Python random, PHP Random) ──
#[test] fn random_new() { compile_ok("var rng = Random();"); }

// ── Stopwatch (same as VB Stopwatch, PHP Stopwatch) ──
#[test] fn stopwatch() { compile_ok("var sw = Stopwatch();"); }

// ── RegExp (same as JS RegExp, Python re, PHP preg_*) ──
#[test] fn regexp_new() { compile_ok("var re = RegExp(r'\\d+');"); }

// ── Sockets (same as VB TcpClient, Python socket, JS net, PHP fsockopen) ──
#[test] fn socket_connect() { compile_ok("var sock = Socket('localhost', 80);"); }
#[test] fn server_socket() { compile_ok("var server = ServerSocket(8080);"); }

// ── Crypto (same as VB SHA256, Python hashlib, JS crypto, PHP md5) ──
#[test] fn crypto_sha256() { compile_ok("var hash = sha256.convert([1, 2, 3]);"); }

// ── Path operations (same as VB Path, Python os.path, JS path, PHP pathinfo) ──
#[test] fn path_join() { compile_ok("var p = Path.join('/tmp', 'test.txt');"); }
#[test] fn path_dirname() { compile_ok("var d = Path.dirname('/tmp/test.txt');"); }
#[test] fn path_basename() { compile_ok("var b = Path.basename('/tmp/test.txt');"); }

// ── Process (same as VB Process, Python subprocess, JS child_process, PHP exec) ──
#[test] fn process_run() { compile_ok("var result = Process.run('ls', ['-la']);"); }

// ── Isolate threading (same as VB Thread, Python threading, JS Worker) ──
#[test] fn isolate_spawn() { compile_ok("void worker(msg) {} Isolate.spawn(worker, 'hello');"); }

// ── DNS (same as VB Dns, Python socket.gethostbyname, JS dns, PHP gethostbyname) ──
#[test] fn dns_lookup() { compile_ok("var addr = InternetAddress.lookup('example.com');"); }

// ── XML (same as VB XDocument, Python xml, JS xml2js, PHP simplexml) ──
#[test] fn xml_parse() { compile_ok("var doc = XmlDocument.parse('<root/>');"); }

// ── Real-world Dart patterns ────────────────────────────────
#[test]
fn http_client() { compile_ok(r#"
var response = http.get('https://api.example.com/data');
"#); }

#[test]
fn file_io() { compile_ok(r#"
var content = File.readAsStringSync('data.txt');
File.writeAsStringSync('output.txt', content);
"#); }

#[test]
fn math_ops() { compile_ok(r#"
var x = math.sqrt(16);
var y = math.pow(2, 10);
var z = math.sin(3.14);
"#); }

#[test]
fn json_roundtrip() { compile_ok(r#"
var data = json.decode('{"name": "Alice"}');
var str = json.encode(data);
"#); }
