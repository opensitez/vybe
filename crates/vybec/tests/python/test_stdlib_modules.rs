use vybec::parser_python::parse;
use vybec::compiler_python::Compiler;

fn compile_ok(src: &str) {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&module);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
}

// ── socket module ───────────────────────────────────────────
#[test] fn socket_connect() { compile_ok("import socket\ns = socket.create_connection(('localhost', 80))\n"); }
#[test] fn socket_dns() { compile_ok("import socket\nip = socket.gethostbyname('example.com')\n"); }
#[test] fn socket_methods() { compile_ok("import socket\ns = socket.socket()\ns.send('hello')\ndata = s.recv(1024)\ns.close()\n"); }
#[test] fn socket_server() { compile_ok("import socket\ns = socket.socket()\ns.bind(('0.0.0.0', 8080))\ns.listen()\nconn = s.accept()\n"); }

// ── sqlite3 module ──────────────────────────────────────────
#[test] fn sqlite3_connect() { compile_ok("import sqlite3\nconn = sqlite3.connect('test.db')\n"); }
#[test] fn sqlite3_cursor() { compile_ok("import sqlite3\nconn = sqlite3.connect('test.db')\ncur = conn.cursor()\ncur.execute('SELECT 1')\nrows = cur.fetchall()\n"); }
#[test] fn sqlite3_fetchone() { compile_ok("import sqlite3\nconn = sqlite3.connect(':memory:')\ncur = conn.cursor()\ncur.execute('CREATE TABLE t (x INT)')\nrow = cur.fetchone()\n"); }

// ── hashlib module ──────────────────────────────────────────
#[test] fn hashlib_sha256() { compile_ok("import hashlib\nh = hashlib.sha256()\n"); }
#[test] fn hashlib_md5() { compile_ok("import hashlib\nh = hashlib.md5()\n"); }

// ── datetime module ─────────────────────────────────────────
#[test] fn datetime_now() { compile_ok("import datetime\nnow = datetime.now()\n"); }
#[test] fn datetime_today() { compile_ok("import datetime\ntoday = datetime.today()\n"); }
#[test] fn datetime_strftime() { compile_ok("import datetime\nnow = datetime.now()\ns = now.strftime('%Y-%m-%d')\n"); }

// ── time module ─────────────────────────────────────────────
#[test] fn time_time() { compile_ok("import time\nt = time.time()\n"); }
#[test] fn time_sleep() { compile_ok("import time\ntime.sleep(1)\n"); }
#[test] fn time_perf_counter() { compile_ok("import time\nt = time.perf_counter()\n"); }

// ── requests / http module ──────────────────────────────────
#[test] fn requests_get() { compile_ok("import requests\nr = requests.get('https://example.com')\n"); }
#[test] fn requests_post() { compile_ok("import requests\nr = requests.post('https://example.com')\n"); }
#[test] fn urllib_urlopen() { compile_ok("import urllib\nr = urllib.urlopen('https://example.com')\n"); }

// ── collections module ──────────────────────────────────────
#[test] fn collections_deque() { compile_ok("import collections\nd = collections.deque()\n"); }
#[test] fn collections_ordered_dict() { compile_ok("import collections\nd = collections.OrderedDict()\n"); }
#[test] fn collections_counter() { compile_ok("import collections\nc = collections.Counter()\n"); }

// ── xml module ──────────────────────────────────────────────
#[test] fn xml_parse() { compile_ok("import xml\ntree = xml.parse('<root/>')\n"); }

// ── os module (expanded) ────────────────────────────────────
#[test] fn os_getcwd() { compile_ok("import os\ncwd = os.getcwd()\n"); }
#[test] fn os_listdir() { compile_ok("import os\nfiles = os.listdir('.')\n"); }
#[test] fn os_mkdir() { compile_ok("import os\nos.mkdir('/tmp/testdir')\n"); }
#[test] fn os_remove() { compile_ok("import os\nos.remove('/tmp/test.txt')\n"); }
#[test] fn os_rename() { compile_ok("import os\nos.rename('old.txt', 'new.txt')\n"); }
#[test] fn os_getenv() { compile_ok("import os\npath = os.getenv('PATH')\n"); }
#[test] fn os_system() { compile_ok("import os\nos.system('ls')\n"); }

// ── os.path module ──────────────────────────────────────────
#[test] fn os_path_exists() { compile_ok("import os\nx = os.path.exists('/tmp')\n"); }
#[test] fn os_path_isfile() { compile_ok("import os\nx = os.path.isfile('/tmp/test.txt')\n"); }
#[test] fn os_path_isdir() { compile_ok("import os\nx = os.path.isdir('/tmp')\n"); }
#[test] fn os_path_join() { compile_ok("import os\np = os.path.join('/tmp', 'test.txt')\n"); }
#[test] fn os_path_dirname() { compile_ok("import os\nd = os.path.dirname('/tmp/test.txt')\n"); }
#[test] fn os_path_basename() { compile_ok("import os\nb = os.path.basename('/tmp/test.txt')\n"); }
#[test] fn os_path_abspath() { compile_ok("import os\np = os.path.abspath('test.txt')\n"); }
#[test] fn os_path_getsize() { compile_ok("import os\ns = os.path.getsize('test.txt')\n"); }

// ── open() with file handle methods ─────────────────────────
#[test] fn open_read() { compile_ok("f = open('test.txt', 'r')\ndata = f.read()\nf.close()\n"); }
#[test] fn open_write() { compile_ok("f = open('test.txt', 'w')\nf.write('hello')\nf.close()\n"); }
#[test] fn open_readline() { compile_ok("f = open('test.txt')\nline = f.readline()\nf.close()\n"); }
#[test] fn open_readlines() { compile_ok("f = open('test.txt')\nlines = f.readlines()\nf.close()\n"); }

// ── Real-world patterns ─────────────────────────────────────
#[test]
fn database_crud() { compile_ok(r#"
import sqlite3
conn = sqlite3.connect('app.db')
cur = conn.cursor()
cur.execute('CREATE TABLE users (id INT, name TEXT)')
cur.execute('INSERT INTO users VALUES (1, "Alice")')
rows = cur.fetchall()
conn.close()
"#); }

#[test]
fn http_json_api() { compile_ok(r#"
import requests
import json
response = requests.get('https://api.example.com/data')
data = json.loads(response)
print(data)
"#); }

#[test]
fn file_processing() { compile_ok(r#"
import os
files = os.listdir('.')
for f in files:
    if os.path.isfile(f):
        size = os.path.getsize(f)
        print(f, size)
"#); }

#[test]
fn timing() { compile_ok(r#"
import time
start = time.perf_counter()
time.sleep(1)
elapsed = time.perf_counter()
print(elapsed)
"#); }
