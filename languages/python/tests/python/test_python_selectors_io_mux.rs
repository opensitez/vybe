use super::helpers::run_python;

// selectors — DefaultSelector, register, select, key, events, SelectorKey

#[test]
fn test_selectors_default_selector_creates_ok() {
    let out = run_python(r#"
import selectors
sel = selectors.DefaultSelector()
print(type(sel).__name__ in ["EpollSelector", "KqueueSelector", "SelectSelector", "PollSelector", "DevpollSelector"])
sel.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_register_returns_selectorkey() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
key = sel.register(sock, selectors.EVENT_READ)
print(type(key).__name__)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["SelectorKey"]);
}

#[test]
fn test_selectors_selectorkey_fields() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
data = {"tag": "my_socket"}
key = sel.register(sock, selectors.EVENT_READ | selectors.EVENT_WRITE, data=data)
print(key.events == (selectors.EVENT_READ | selectors.EVENT_WRITE))
print(key.data)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["True", "{'tag': 'my_socket'}"]);
}

#[test]
fn test_selectors_unregister_removes_key() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ)
sel.unregister(sock)
try:
    sel.unregister(sock)
    print("no error")
except KeyError:
    print("KeyError")
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["KeyError"]);
}

#[test]
fn test_selectors_get_key() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ, data=42)
key = sel.get_key(sock)
print(key.data)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_selectors_get_map_lists_registered() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
socks = [socket.socket() for _ in range(3)]
for s in socks:
    s.setblocking(False)
    sel.register(s, selectors.EVENT_READ)
print(len(sel.get_map()) == 3)
sel.close()
for s in socks: s.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_event_read_constant() {
    let out = run_python(r#"
import selectors
print(selectors.EVENT_READ)
print(selectors.EVENT_WRITE)
"#);
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_selectors_select_timeout_returns_immediately() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
# No registered sockets — select with timeout=0 should return []
result = sel.select(timeout=0)
print(result)
sel.close()
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_selectors_select_writable_socket() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
# A newly created connected socket is usually writable immediately
client = socket.socket()
client.setblocking(False)
try:
    client.connect(("127.0.0.1", 9))  # port 9 = discard (connection refused)
except OSError:
    pass
sel.register(client, selectors.EVENT_WRITE)
ready = sel.select(timeout=0.1)
# We just check the type is correct — no error
print(isinstance(ready, list))
sel.close()
client.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_modify_changes_events() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ)
new_key = sel.modify(sock, selectors.EVENT_WRITE)
print(new_key.events == selectors.EVENT_WRITE)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_modify_changes_data() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ, data="old")
new_key = sel.modify(sock, selectors.EVENT_READ, data="new")
print(new_key.data)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["new"]);
}

#[test]
fn test_selectors_close_allows_reuse() {
    let out = run_python(r#"
import selectors
sel = selectors.DefaultSelector()
sel.close()
# After close, creating a new one must work
sel2 = selectors.DefaultSelector()
print("ok")
sel2.close()
"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_selectors_context_manager() {
    let out = run_python(r#"
import selectors
with selectors.DefaultSelector() as sel:
    print(type(sel).__name__ != "")
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_epoll_selector_available_on_linux() {
    let out = run_python(r#"
import selectors, sys
if sys.platform == "linux":
    print(hasattr(selectors, "EpollSelector"))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_kqueue_selector_available_on_bsd() {
    let out = run_python(r#"
import selectors, sys
if sys.platform in ("darwin", "freebsd"):
    print(hasattr(selectors, "KqueueSelector"))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_multiple_register_same_fd_raises() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ)
try:
    sel.register(sock, selectors.EVENT_WRITE)
    print("no error")
except KeyError:
    print("KeyError")
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["KeyError"]);
}

#[test]
fn test_selectors_selectorkey_fileobj_matches() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
key = sel.register(sock, selectors.EVENT_READ)
print(key.fileobj is sock)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_selectorkey_fd_is_int() {
    let out = run_python(r#"
import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
key = sel.register(sock, selectors.EVENT_READ)
print(isinstance(key.fd, int))
print(key.fd >= 0)
sel.close()
sock.close()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_selectors_select_returns_list_of_tuples() {
    let out = run_python(r#"
import selectors
sel = selectors.DefaultSelector()
result = sel.select(timeout=0)
print(isinstance(result, list))
sel.close()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_selectors_get_map_empty_when_none_registered() {
    let out = run_python(r#"
import selectors
sel = selectors.DefaultSelector()
print(len(sel.get_map()) == 0)
sel.close()
"#);
    assert_eq!(out, vec!["True"]);
}
