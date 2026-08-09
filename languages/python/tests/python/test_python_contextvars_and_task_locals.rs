// Python: contextvars — ContextVar, Token, copy_context, Context execution, and state isolation
use super::helpers::run_python;

#[test]
fn test_contextvar_default_value() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=42)
print(var.get())
"##;
    assert_eq!(run_python(script), vec!["42"]);
}

#[test]
fn test_contextvar_no_default_raises_lookup_error() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var')
try:
    var.get()
except LookupError:
    print('LOOKUP_ERROR')
"##;
    assert_eq!(run_python(script), vec!["LOOKUP_ERROR"]);
}

#[test]
fn test_contextvar_get_with_explicit_default() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var')
print(var.get('fallback'))
"##;
    assert_eq!(run_python(script), vec!["fallback"]);
}

#[test]
fn test_contextvar_set_and_get() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default='initial')
var.set('updated')
print(var.get())
"##;
    assert_eq!(run_python(script), vec!["updated"]);
}

#[test]
fn test_contextvar_token_reset() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=10)
tok = var.set(20)
print(var.get())
var.reset(tok)
print(var.get())
"##;
    assert_eq!(run_python(script), vec!["20", "10"]);
}

#[test]
fn test_contextvar_multiple_set_reset_chain() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=0)
t1 = var.set(1)
t2 = var.set(2)
t3 = var.set(3)
print(var.get())
var.reset(t3)
print(var.get())
var.reset(t2)
print(var.get())
var.reset(t1)
print(var.get())
"##;
    assert_eq!(run_python(script), vec!["3", "2", "1", "0"]);
}

#[test]
fn test_contextvar_token_var_attribute() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var_name', default=100)
tok = var.set(200)
print(tok.var is var)
print(tok.var.name)
"##;
    assert_eq!(run_python(script), vec!["True", "var_name"]);
}

#[test]
fn test_contextvar_copy_context_isolation() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default='base')
var.set('parent_val')

ctx = contextvars.copy_context()

def child_func():
    var.set('child_val')
    return var.get()

res = ctx.run(child_func)
print('child:', res)
print('parent:', var.get())
"##;
    assert_eq!(
        run_python(script),
        vec!["child: child_val", "parent: parent_val"]
    );
}

#[test]
fn test_contextvar_copy_context_keys_values() {
    let script = r##"
import contextvars

v1 = contextvars.ContextVar('v1')
v2 = contextvars.ContextVar('v2')

v1.set(100)
v2.set('hello')

ctx = contextvars.copy_context()
items = {k.name: v for k, v in ctx.items()}
print(items['v1'], items['v2'])
"##;
    assert_eq!(run_python(script), vec!["100 hello"]);
}

#[test]
fn test_contextvar_copy_context_get_method() {
    let script = r##"
import contextvars

v1 = contextvars.ContextVar('v1')
v1.set('val1')

ctx = contextvars.copy_context()
print(ctx.get(v1))
"##;
    assert_eq!(run_python(script), vec!["val1"]);
}

#[test]
fn test_contextvar_copy_context_len() {
    let script = r##"
import contextvars

v1 = contextvars.ContextVar('v1')
v2 = contextvars.ContextVar('v2')
v1.set(1)
v2.set(2)

ctx = contextvars.copy_context()
print(len(ctx) >= 2)
"##;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_contextvar_token_old_value_empty() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var')
tok = var.set('first')
print(tok.old_value is contextvars.Token.MISSING)
"##;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_contextvar_token_old_value_present() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var')
var.set('first')
tok2 = var.set('second')
print(tok2.old_value)
"##;
    assert_eq!(run_python(script), vec!["first"]);
}

#[test]
fn test_contextvar_in_generator() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=0)

def gen():
    var.set(10)
    yield var.get()
    var.set(20)
    yield var.get()

g = gen()
print(next(g))
print(next(g))
print(var.get())
"##;
    assert_eq!(run_python(script), vec!["10", "20", "20"]);
}

#[test]
fn test_contextvar_threads_isolation() {
    let script = r##"
import contextvars
import threading

var = contextvars.ContextVar('var', default='main')

def worker(val):
    var.set(val)
    print(threading.current_thread().name, var.get())

var.set('parent')
t = threading.Thread(target=worker, args=('thread_val',), name='T1')
t.start()
t.join()
print('main:', var.get())
"##;
    assert_eq!(run_python(script), vec!["T1 thread_val", "main: parent"]);
}

#[test]
fn test_contextvar_context_run_args_kwargs() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=0)
ctx = contextvars.copy_context()

def add(a, b, extra=0):
    var.set(a + b + extra)
    return var.get()

res = ctx.run(add, 5, 10, extra=2)
print(res)
"##;
    assert_eq!(run_python(script), vec!["17"]);
}

#[test]
fn test_contextvar_token_used_twice_raises_runtime_error() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default=0)
tok = var.set(10)
var.reset(tok)
try:
    var.reset(tok)
except RuntimeError:
    print('TOKEN_REUSED_ERROR')
"##;
    assert_eq!(run_python(script), vec!["TOKEN_REUSED_ERROR"]);
}

#[test]
fn test_contextvar_token_reset_wrong_var_raises_value_error() {
    let script = r##"
import contextvars

v1 = contextvars.ContextVar('v1')
v2 = contextvars.ContextVar('v2')

tok = v1.set('a')
try:
    v2.reset(tok)
except ValueError:
    print('WRONG_VAR_TOKEN')
"##;
    assert_eq!(run_python(script), vec!["WRONG_VAR_TOKEN"]);
}

#[test]
fn test_contextvar_context_type_check() {
    let script = r##"
import contextvars

ctx1 = contextvars.copy_context()
print(type(ctx1).__name__)
"##;
    assert_eq!(run_python(script), vec!["Context"]);
}

#[test]
fn test_contextvar_context_run_exception_propagation() {
    let script = r##"
import contextvars

var = contextvars.ContextVar('var', default='ok')
ctx = contextvars.copy_context()

def failing():
    var.set('modified')
    raise ValueError('custom failure')

try:
    ctx.run(failing)
except ValueError as e:
    print('CAUGHT:', e)
"##;
    assert_eq!(run_python(script), vec!["CAUGHT: custom failure"]);
}
