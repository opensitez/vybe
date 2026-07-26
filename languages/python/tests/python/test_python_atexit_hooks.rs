// Python atexit module — register, unregister, call order
use super::helpers::run_python;

#[test]
fn test_atexit_register_called() {
    let script = r#"
import atexit

results = []

def cleanup():
    results.append("cleaned")

atexit.register(cleanup)
# manually invoke since we can't wait for interpreter exit
atexit._run_exitfuncs()
print(results)
"#;
    assert_eq!(run_python(script), vec!["['cleaned']"]);
}

#[test]
fn test_atexit_lifo_order() {
    let script = r#"
import atexit

order = []

def first():
    order.append(1)

def second():
    order.append(2)

atexit.register(first)
atexit.register(second)
atexit._run_exitfuncs()
print(order)
"#;
    assert_eq!(run_python(script), vec!["[2, 1]"]);
}

#[test]
fn test_atexit_register_with_args() {
    let script = r#"
import atexit

log = []

def record(msg, count):
    for _ in range(count):
        log.append(msg)

atexit.register(record, "bye", 3)
atexit._run_exitfuncs()
print(log)
"#;
    assert_eq!(run_python(script), vec!["['bye', 'bye', 'bye']"]);
}

#[test]
fn test_atexit_unregister() {
    let script = r#"
import atexit

called = []

def never():
    called.append("never")

atexit.register(never)
atexit.unregister(never)
atexit._run_exitfuncs()
print(called)
"#;
    assert_eq!(run_python(script), vec!["[]"]);
}

#[test]
fn test_atexit_register_same_func_twice() {
    let script = r#"
import atexit

count = []

def inc():
    count.append(1)

atexit.register(inc)
atexit.register(inc)
atexit._run_exitfuncs()
print(sum(count))
"#;
    assert_eq!(run_python(script), vec!["2"]);
}
