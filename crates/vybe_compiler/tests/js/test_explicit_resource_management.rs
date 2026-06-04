/// Explicit Resource Management (ES2025) — using declarations, Symbol.dispose,
/// Symbol.asyncDispose, DisposableStack, AsyncDisposableStack, await using.
use super::helpers::run_js;

// ── Symbol.dispose ────────────────────────────────────────────────────────────

#[test]
fn symbol_dispose_exists() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof Symbol.dispose);
console.log(typeof Symbol.asyncDispose);
"#
        ),
        vec!["symbol", "symbol"]
    );
}

#[test]
fn object_with_dispose_method() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const res = {
    [Symbol.dispose]() { log.push("disposed"); }
};
{
    using r = res;
}
console.log(log.join(","));
"#
        ),
        vec!["disposed"]
    );
}

#[test]
fn using_disposes_on_block_exit() {
    assert_eq!(
        run_js(
            r#"
const log = [];
function makeResource(name) {
    return { [Symbol.dispose]() { log.push("close:" + name); } };
}
{
    using a = makeResource("A");
    using b = makeResource("B");
    log.push("work");
}
console.log(log.join(","));
"#
        ),
        vec!["work,close:B,close:A"]
    );
}

#[test]
fn using_disposes_in_lifo_order() {
    assert_eq!(
        run_js(
            r#"
const order = [];
{
    using r1 = { [Symbol.dispose]() { order.push(1); } };
    using r2 = { [Symbol.dispose]() { order.push(2); } };
    using r3 = { [Symbol.dispose]() { order.push(3); } };
}
console.log(order.join(","));
"#
        ),
        vec!["3,2,1"]
    );
}

#[test]
fn using_disposes_even_on_throw() {
    assert_eq!(
        run_js(
            r#"
const log = [];
try {
    using r = { [Symbol.dispose]() { log.push("disposed"); } };
    throw new Error("oops");
} catch (e) {
    log.push("caught:" + e.message);
}
console.log(log.join(","));
"#
        ),
        vec!["disposed,caught:oops"]
    );
}

#[test]
fn using_null_or_undefined_is_allowed() {
    assert_eq!(
        run_js(
            r#"
let ok = true;
try {
    using r = null;
    using s = undefined;
} catch {
    ok = false;
}
console.log(ok);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn using_non_disposable_object_throws() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    using r = { value: 42 }; // no Symbol.dispose
} catch (e) {
    threw = e instanceof TypeError;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

// ── using in for loop ─────────────────────────────────────────────────────────

#[test]
fn using_in_for_of_disposes_each_iteration() {
    assert_eq!(
        run_js(
            r#"
const log = [];
function makeRes(n) {
    return { n, [Symbol.dispose]() { log.push("d" + this.n); } };
}
for (using r of [makeRes(1), makeRes(2), makeRes(3)]) {
    log.push("u" + r.n);
}
console.log(log.join(","));
"#
        ),
        vec!["u1,d1,u2,d2,u3,d3"]
    );
}

// ── DisposableStack ───────────────────────────────────────────────────────────

#[test]
fn disposable_stack_basic_use() {
    assert_eq!(
        run_js(
            r#"
const log = [];
{
    using stack = new DisposableStack();
    stack.defer(() => log.push("cleanup1"));
    stack.defer(() => log.push("cleanup2"));
    log.push("work");
}
console.log(log.join(","));
"#
        ),
        vec!["work,cleanup2,cleanup1"]
    );
}

#[test]
fn disposable_stack_adopt() {
    assert_eq!(
        run_js(
            r#"
const log = [];
{
    using stack = new DisposableStack();
    const handle = stack.adopt({ id: 1 }, (h) => log.push("close:" + h.id));
    log.push("use:" + handle.id);
}
console.log(log.join(","));
"#
        ),
        vec!["use:1,close:1"]
    );
}

#[test]
fn disposable_stack_use() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const res = { [Symbol.dispose]() { log.push("disposed"); } };
{
    using stack = new DisposableStack();
    stack.use(res);
    log.push("work");
}
console.log(log.join(","));
"#
        ),
        vec!["work,disposed"]
    );
}

#[test]
fn disposable_stack_disposed_property() {
    assert_eq!(
        run_js(
            r#"
const stack = new DisposableStack();
console.log(stack.disposed);
stack.dispose();
console.log(stack.disposed);
"#
        ),
        vec!["false", "true"]
    );
}

#[test]
fn disposable_stack_move_transfers_ownership() {
    assert_eq!(
        run_js(
            r#"
const log = [];
let outer;
{
    using stack = new DisposableStack();
    stack.defer(() => log.push("cleanup"));
    outer = stack.move();
    log.push("inner disposed:" + stack.disposed);
}
log.push("outer disposed before:" + outer.disposed);
outer.dispose();
log.push("outer disposed after:" + outer.disposed);
console.log(log.join(","));
"#
        ),
        vec!["inner disposed:true,outer disposed before:false,cleanup,outer disposed after:true"]
    );
}

// ── await using ───────────────────────────────────────────────────────────────

#[test]
fn await_using_calls_async_dispose() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function main() {
    await using r = {
        [Symbol.asyncDispose]() {
            return Promise.resolve().then(() => log.push("async disposed"));
        }
    };
    log.push("work");
}
main().then(() => console.log(log.join(",")));
"#
        ),
        vec!["work,async disposed"]
    );
}

#[test]
fn await_using_falls_back_to_sync_dispose() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function main() {
    await using r = { [Symbol.dispose]() { log.push("sync disposed"); } };
    log.push("work");
}
main().then(() => console.log(log.join(",")));
"#
        ),
        vec!["work,sync disposed"]
    );
}

// ── AsyncDisposableStack ──────────────────────────────────────────────────────

#[test]
fn async_disposable_stack_basic() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function main() {
    await using stack = new AsyncDisposableStack();
    stack.defer(async () => {
        await Promise.resolve();
        log.push("cleanup");
    });
    log.push("work");
}
main().then(() => console.log(log.join(",")));
"#
        ),
        vec!["work,cleanup"]
    );
}
