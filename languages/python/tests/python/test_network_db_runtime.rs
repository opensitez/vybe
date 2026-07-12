//! socket, sqlite3, ssl, http.client — runtime where possible.

crate::runtime_case!(socket_module, "import socket\nprint(socket.AF_INET)\n", "2");
crate::runtime_case!(
    socket_sock_stream,
    "import socket\nprint(socket.SOCK_STREAM)\n",
    "1"
);
crate::runtime_case!(
    socket_create_connection_exists,
    "import socket\nprint(callable(socket.create_connection))\n",
    "True"
);
crate::runtime_case!(
    socket_socket_class,
    "import socket\nprint(callable(socket.socket))\n",
    "True"
);
crate::runtime_case!(
    socket_inet_aton,
    "import socket\nprint(socket.inet_aton('127.0.0.1'))\n",
    "b'\\x7f\\x00\\x00\\x01'"
);
crate::runtime_case!(
    socket_inet_ntoa,
    "import socket\nprint(socket.inet_ntoa(b'\\x7f\\x00\\x00\\x01'))\n",
    "127.0.0.1"
);
crate::runtime_case!(
    socket_gethostname,
    "import socket\nprint(isinstance(socket.gethostname(), str))\n",
    "True"
);
crate::runtime_case!(
    socket_getaddrinfo_exists,
    "import socket\nprint(callable(socket.getaddrinfo))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_connect_memory,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nprint(conn is not None)\n",
    "True"
);
crate::runtime_case!(
    sqlite3_execute_select,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\ncur = conn.cursor()\ncur.execute('SELECT 1')\nprint(cur.fetchone()[0])\n",
    "1"
);
crate::runtime_case!(
    sqlite3_create_table,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.execute('CREATE TABLE t (x INTEGER)')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    sqlite3_insert_fetch,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.execute('CREATE TABLE t (x INTEGER)')\nconn.execute('INSERT INTO t VALUES (42)')\nprint(conn.execute('SELECT x FROM t').fetchone()[0])\n",
    "42"
);
crate::runtime_case!(
    sqlite3_rowcount,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\ncur = conn.cursor()\ncur.execute('SELECT 1')\nprint(cur.rowcount)\n",
    "-1"
);
crate::runtime_case!(
    sqlite3_paramstyle,
    "import sqlite3\nprint(sqlite3.paramstyle)\n",
    "qmark"
);
crate::runtime_case!(
    sqlite3_version,
    "import sqlite3\nprint(isinstance(sqlite3.sqlite_version, str))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_exception,
    "import sqlite3\ntry:\n sqlite3.connect(':memory:').execute('INVALID SQL')\n print('ok')\nexcept sqlite3.Error:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    ssl_create_default_context,
    "import ssl\nctx = ssl.create_default_context()\nprint(ctx is not None)\n",
    "True"
);
crate::runtime_case!(
    ssl_protocol_tls,
    "import ssl\nprint(hasattr(ssl, 'PROTOCOL_TLS'))\n",
    "True"
);
crate::runtime_case!(ssl_cert_none, "import ssl\nprint(ssl.CERT_NONE)\n", "0");
crate::runtime_case!(
    ssl_wrap_socket_exists,
    "import ssl\nprint(callable(ssl.wrap_socket))\n",
    "True"
);
crate::runtime_case!(
    http_client_ok,
    "import http.client\nprint(http.client.OK)\n",
    "200"
);
crate::runtime_case!(
    http_client_not_found,
    "import http.client\nprint(http.client.NOT_FOUND)\n",
    "404"
);
crate::runtime_case!(
    http_client_responses,
    "import http.client\nprint(200 in http.client.responses)\n",
    "True"
);
crate::runtime_case!(
    http_client_httpconnection,
    "import http.client\nprint(callable(http.client.HTTPConnection))\n",
    "True"
);
crate::runtime_case!(
    http_client_httpsconnection,
    "import http.client\nprint(callable(http.client.HTTPSConnection))\n",
    "True"
);
crate::runtime_case!(
    socket_timeout_default,
    "import socket\nprint(socket.getdefaulttimeout() is None)\n",
    "True"
);
crate::runtime_case!(
    socket_error_hierarchy,
    "import socket\nprint(issubclass(socket.timeout, OSError))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_thread_safety,
    "import sqlite3\nprint(sqlite3.threadsafety >= 0)\n",
    "True"
);
crate::runtime_case!(
    sqlite3_row_factory,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.row_factory = sqlite3.Row\nprint(callable(conn.row_factory))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_isolation_level,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nprint(conn.isolation_level is None or isinstance(conn.isolation_level, str))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_commit,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.commit()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    sqlite3_close,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.close()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    ssl_has_tls_version,
    "import ssl\nprint(hasattr(ssl, 'TLSVersion'))\n",
    "True"
);
crate::runtime_case!(
    http_client_badstatus,
    "import http.client\nprint(issubclass(http.client.BadStatusLine, Exception))\n",
    "True"
);
crate::runtime_case!(
    socket_inet_pton,
    "import socket\nprint(socket.inet_pton(socket.AF_INET, '127.0.0.1'))\n",
    "b'\\x7f\\x00\\x00\\x01'"
);
crate::runtime_case!(
    socket_ntohl,
    "import socket\nprint(socket.ntohl(0x7f000001))\n",
    "2130706433"
);
crate::runtime_case!(
    sqlite3_many_execute,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nconn.execute('CREATE TABLE t (x INTEGER)')\nconn.executemany('INSERT INTO t VALUES (?)', [(1,), (2,)])\nprint(conn.execute('SELECT COUNT(*) FROM t').fetchone()[0])\n",
    "2"
);
crate::runtime_case!(
    sqlite3_pragma,
    "import sqlite3\nconn = sqlite3.connect(':memory:')\nprint(conn.execute('PRAGMA user_version').fetchone()[0])\n",
    "0"
);
crate::runtime_case!(
    ssl_match_hostname,
    "import ssl\nprint(callable(ssl.match_hostname))\n",
    "True"
);
crate::runtime_case!(
    http_client_parse_headers,
    "import http.client\nprint(callable(http.client.parse_headers))\n",
    "True"
);
crate::runtime_case!(
    socket_dup,
    "import socket\nprint(hasattr(socket.socket, 'dup'))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_binary,
    "import sqlite3\nprint(sqlite3.Binary(b'hi'))\n",
    "b'hi'"
);
crate::runtime_case!(
    ssl_enum_cert,
    "import ssl\nprint(hasattr(ssl, 'Purpose'))\n",
    "True"
);
crate::runtime_case!(
    http_client_immutable,
    "import http.client\nprint(http.client.responses[404])\n",
    "Not Found"
);
crate::runtime_case!(
    socket_has_ipv6,
    "import socket\nprint(hasattr(socket, 'AF_INET6'))\n",
    "True"
);
crate::runtime_case!(
    sqlite3_adapters,
    "import sqlite3\nprint(hasattr(sqlite3, 'register_adapter'))\n",
    "True"
);
crate::runtime_case!(
    ssl_default_ciphers,
    "import ssl\nprint(isinstance(ssl.create_default_context().get_ciphers(), list))\n",
    "True"
);

crate::compile_case!(
    socket_connect_ex,
    "import socket\ns = socket.socket()\ns.connect_ex(('127.0.0.1', 9))\n"
);
crate::compile_case!(
    sqlite3_backup,
    "import sqlite3\nsrc = sqlite3.connect(':memory:')\n"
);
crate::compile_case!(ssl_sslsocket, "import ssl\nssl.SSLSocket\n");
crate::compile_case!(
    http_client_request,
    "import http.client\nhttp.client.HTTPConnection('localhost')\n"
);
crate::compile_case!(
    socketserver_tcpserver,
    "import socketserver\nsocketserver.TCPServer\n"
);
