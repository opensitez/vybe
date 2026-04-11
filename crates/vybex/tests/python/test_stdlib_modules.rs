use super::helpers::compile_ok;

// socket module
#[test] fn socket_connect() { compile_ok("import socket\ns = socket.create_connection(('localhost', 80))\n"); }
#[test] fn socket_methods() { compile_ok("import socket\ns = socket.socket()\ns.send('hello')\ndata = s.recv(1024)\ns.close()\n"); }

// sqlite3 module
#[test] fn sqlite3_connect() { compile_ok("import sqlite3\nconn = sqlite3.connect('test.db')\n"); }
#[test] fn sqlite3_cursor() { compile_ok("import sqlite3\nconn = sqlite3.connect('test.db')\ncur = conn.cursor()\ncur.execute('SELECT 1')\nrows = cur.fetchall()\n"); }

// hashlib module
#[test] fn hashlib_sha256() { compile_ok("import hashlib\nh = hashlib.sha256()\n"); }
#[test] fn hashlib_md5() { compile_ok("import hashlib\nh = hashlib.md5()\n"); }

// datetime module
#[test] fn datetime_now() { compile_ok("import datetime\nnow = datetime.now()\n"); }
#[test] fn datetime_strftime() { compile_ok("import datetime\nnow = datetime.now()\ns = now.strftime('%Y-%m-%d')\n"); }

// time module
#[test] fn time_perf_counter() { compile_ok("import time\nt = time.perf_counter()\n"); }

// requests module
#[test] fn requests_get() { compile_ok("import requests\nr = requests.get('https://example.com')\n"); }
#[test] fn requests_post() { compile_ok("import requests\nr = requests.post('https://example.com')\n"); }

// collections module
#[test] fn collections_deque() { compile_ok("import collections\nd = collections.deque()\n"); }
#[test] fn collections_ordered_dict() { compile_ok("import collections\nd = collections.OrderedDict()\n"); }
#[test] fn collections_counter() { compile_ok("import collections\nc = collections.Counter()\n"); }

// xml module
#[test] fn xml_parse() { compile_ok("import xml\ntree = xml.parse('<root/>')\n"); }
