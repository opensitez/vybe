use super::helpers::compile_ok;

// Match/case

#[test]
fn match_basic_value() {
    compile_ok(
        "x = 42\nmatch x:\n    case 1:\n        print('one')\n    case 42:\n        print('forty-two')\n    case _:\n        print('other')\n",
    );
}

#[test]
fn match_wildcard() {
    compile_ok("match 'hello':\n    case _:\n        print('anything')\n");
}

#[test]
fn match_or_pattern() {
    compile_ok(
        "x = 2\nmatch x:\n    case 1 | 2 | 3:\n        print('small')\n    case _:\n        print('big')\n",
    );
}

#[test]
fn match_with_guard() {
    compile_ok(
        "x = 10\nmatch x:\n    case n if n > 5:\n        print('big')\n    case _:\n        print('small')\n",
    );
}

#[test]
fn match_none() {
    compile_ok(
        "x = None\nmatch x:\n    case None:\n        print('none')\n    case _:\n        print('other')\n",
    );
}

#[test]
fn match_string() {
    compile_ok(
        "cmd = 'quit'\nmatch cmd:\n    case 'quit' | 'exit':\n        print('bye')\n    case 'help':\n        print('help')\n    case _:\n        pass\n",
    );
}

// Math module

#[test]
fn math_sqrt() {
    compile_ok("import math\nx = math.sqrt(16)\n");
}
#[test]
fn math_pi() {
    compile_ok("import math\narea = math.pi * r * r\n");
}
#[test]
fn math_sin_cos() {
    compile_ok("import math\nx = math.sin(0.5)\ny = math.cos(0.5)\n");
}
#[test]
fn math_log() {
    compile_ok("import math\nx = math.log(100)\n");
}
#[test]
fn math_floor_ceil() {
    compile_ok("import math\nx = math.floor(3.7)\ny = math.ceil(3.2)\n");
}
#[test]
fn math_isnan() {
    compile_ok("import math\nb = math.isnan(float('nan'))\n");
}
#[test]
fn math_inf() {
    compile_ok("import math\nx = math.inf\n");
}
#[test]
fn math_e() {
    compile_ok("import math\nx = math.e\n");
}

// JSON module

#[test]
fn json_loads() {
    compile_ok("import json\ndata = json.loads('{\"a\": 1}')\n");
}
#[test]
fn json_dumps() {
    compile_ok("import json\ns = json.dumps({'a': 1})\n");
}

// Random module

#[test]
fn random_random() {
    compile_ok("import random\nx = random.random()\n");
}
#[test]
fn random_randint() {
    compile_ok("import random\nn = random.randint(1, 10)\n");
}
#[test]
fn random_choice() {
    compile_ok("import random\nx = random.choice([1, 2, 3])\n");
}

// re module

#[test]
fn re_search() {
    compile_ok("import re\nm = re.search(r'\\d+', 'abc123')\n");
}
#[test]
fn re_findall() {
    compile_ok("import re\nmatches = re.findall(r'\\d+', 'a1b2c3')\n");
}
#[test]
fn re_sub() {
    compile_ok("import re\ns = re.sub(r'\\d', 'X', 'a1b2')\n");
}

// OS module

#[test]
fn os_getcwd() {
    compile_ok("import os\ncwd = os.getcwd()\n");
}
#[test]
fn os_listdir() {
    compile_ok("import os\nfiles = os.listdir('.')\n");
}
#[test]
fn os_mkdir() {
    compile_ok("import os\nos.mkdir('/tmp/testdir')\n");
}
#[test]
fn os_remove() {
    compile_ok("import os\nos.remove('/tmp/test.txt')\n");
}
#[test]
fn os_rename() {
    compile_ok("import os\nos.rename('old.txt', 'new.txt')\n");
}
#[test]
fn os_path_exists() {
    compile_ok("import os\nx = os.path.exists('/tmp')\n");
}
#[test]
fn os_path_isfile() {
    compile_ok("import os\nx = os.path.isfile('/tmp/test.txt')\n");
}
#[test]
fn os_path_isdir() {
    compile_ok("import os\nx = os.path.isdir('/tmp')\n");
}
#[test]
fn os_path_join() {
    compile_ok("import os\np = os.path.join('/tmp', 'test.txt')\n");
}
#[test]
fn os_path_dirname() {
    compile_ok("import os\nd = os.path.dirname('/tmp/test.txt')\n");
}
#[test]
fn os_path_basename() {
    compile_ok("import os\nb = os.path.basename('/tmp/test.txt')\n");
}
#[test]
fn os_path_abspath() {
    compile_ok("import os\np = os.path.abspath('test.txt')\n");
}

// sys module

#[test]
fn sys_exit() {
    compile_ok("import sys\nsys.exit(0)\n");
}

// time module

#[test]
fn time_time() {
    compile_ok("import time\nt = time.time()\n");
}
#[test]
fn time_sleep() {
    compile_ok("import time\ntime.sleep(1)\n");
}

// open() with file methods

#[test]
fn open_read() {
    compile_ok("f = open('test.txt', 'r')\ndata = f.read()\nf.close()\n");
}
#[test]
fn open_write() {
    compile_ok("f = open('test.txt', 'w')\nf.write('hello')\nf.close()\n");
}
#[test]
fn open_readline() {
    compile_ok("f = open('test.txt')\nline = f.readline()\nf.close()\n");
}
#[test]
fn open_readlines() {
    compile_ok("f = open('test.txt')\nlines = f.readlines()\nf.close()\n");
}

// threading module

#[test]
fn threading_thread() {
    compile_ok(
        "import threading\ndef worker():\n    print('hello from thread')\nt = threading.Thread(worker)\n",
    );
}

// Cross-language compat

#[test]
fn class_tostring_alias() {
    compile_ok(
        "class Dog:\n    def __init__(self, name):\n        self.name = name\n    def __str__(self):\n        return self.name\nd = Dog('Rex')\n",
    );
}

// Real-world patterns

#[test]
fn http_json_api() {
    compile_ok(
        r#"
import requests
import json
response = requests.get('https://api.example.com/data')
data = json.loads(response)
print(data)
"#,
    );
}

#[test]
fn file_processing() {
    compile_ok(
        r#"
import os
files = os.listdir('.')
for f in files:
    if os.path.isfile(f):
        size = os.path.getsize(f)
        print(f, size)
"#,
    );
}

#[test]
fn database_crud() {
    compile_ok(
        r#"
import sqlite3
conn = sqlite3.connect('app.db')
cur = conn.cursor()
cur.execute('CREATE TABLE users (id INT, name TEXT)')
cur.execute('INSERT INTO users VALUES (1, "Alice")')
rows = cur.fetchall()
conn.close()
"#,
    );
}
