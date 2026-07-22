use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: queueMicrotask & Job Queue Execution Ordering
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_queue_microtask_basic_execution() {
    let src = r#"
console.log("Sync 1");
queueMicrotask(() => console.log("Microtask 1"));
console.log("Sync 2");
"#;
    assert_eq!(run_js(src), vec!["Sync 1", "Sync 2", "Microtask 1"]);
}

#[test]
fn test_js_queue_microtask_interleaved_with_promise_then() {
    let src = r#"
queueMicrotask(() => console.log("Queue 1"));
Promise.resolve().then(() => console.log("Promise 1"));
queueMicrotask(() => console.log("Queue 2"));
"#;
    assert_eq!(run_js(src), vec!["Queue 1", "Promise 1", "Queue 2"]);
}

#[test]
fn test_js_queue_microtask_nested_microtasks() {
    let src = r#"
queueMicrotask(() => {
    console.log("Level 1");
    queueMicrotask(() => console.log("Level 2"));
});
console.log("Sync End");
"#;
    assert_eq!(run_js(src), vec!["Sync End", "Level 1", "Level 2"]);
}

#[test]
fn test_js_queue_microtask_throws_exception_does_not_block_subsequent() {
    let src = r#"
queueMicrotask(() => {
    throw new Error("Microtask Crash");
});
queueMicrotask(() => {
    console.log("Subsequent Microtask Ran");
});
"#;
    assert_eq!(run_js(src), vec!["Subsequent Microtask Ran"]);
}

#[test]
fn test_js_queue_microtask_fifo_ordering() {
    let src = r#"
for (let i = 1; i <= 3; i++) {
    queueMicrotask(() => console.log("Task " + i));
}
"#;
    assert_eq!(run_js(src), vec!["Task 1", "Task 2", "Task 3"]);
}

#[test]
fn test_js_queue_microtask_non_callable_argument_throws_typeerror() {
    let src = r#"
try {
    queueMicrotask(12345);
} catch (e) {
    console.log("QueueMicrotask Non-Callable TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["QueueMicrotask Non-Callable TypeError"]);
}

#[test]
fn test_js_queue_microtask_with_arguments_binding() {
    let src = r#"
function schedule(val) {
    queueMicrotask(() => console.log("Scheduled: " + val));
}
schedule("A");
schedule("B");
"#;
    assert_eq!(run_js(src), vec!["Scheduled: A", "Scheduled: B"]);
}

#[test]
fn test_js_queue_microtask_this_unbound() {
    let src = r#"
const obj = {
    method() {
        queueMicrotask(function() {
            console.log(this === undefined);
        });
    }
};
obj.method();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_queue_microtask_arrow_function_lexical_this() {
    let src = r#"
const obj = {
    name: "LexicalObj",
    method() {
        queueMicrotask(() => {
            console.log(this.name);
        });
    }
};
obj.method();
"#;
    assert_eq!(run_js(src), vec!["LexicalObj"]);
}

#[test]
fn test_js_queue_microtask_state_mutation() {
    let src = r#"
let state = 0;
queueMicrotask(() => { state += 10; });
queueMicrotask(() => { state *= 2; });
queueMicrotask(() => { console.log("Final State: " + state); });
"#;
    assert_eq!(run_js(src), vec!["Final State: 20"]);
}

#[test]
fn test_js_queue_microtask_inside_promise_then() {
    let src = r#"
Promise.resolve().then(() => {
    console.log("Promise Then");
    queueMicrotask(() => console.log("Nested QueueMicrotask"));
});
"#;
    assert_eq!(run_js(src), vec!["Promise Then", "Nested QueueMicrotask"]);
}

#[test]
fn test_js_queue_microtask_inside_async_function() {
    let src = r#"
async function run() {
    console.log("Async Start");
    queueMicrotask(() => console.log("Queue inside Async"));
    await Promise.resolve();
    console.log("Async After Await");
}
run();
console.log("Main Thread End");
"#;
    assert_eq!(
        run_js(src),
        vec![
            "Async Start",
            "Main Thread End",
            "Queue inside Async",
            "Async After Await"
        ]
    );
}

#[test]
fn test_js_queue_microtask_recursion_drain_queue() {
    let src = r#"
let count = 0;
function recurse() {
    count++;
    if (count < 3) {
        queueMicrotask(recurse);
    }
}
queueMicrotask(recurse);
Promise.resolve().then(() => console.log("Microtask Count: " + count));
"#;
    assert_eq!(run_js(src), vec!["Microtask Count: 3"]);
}

#[test]
fn test_js_queue_microtask_multiple_independent_callbacks() {
    let src = r#"
const fn1 = () => console.log("Fn1");
const fn2 = () => console.log("Fn2");
queueMicrotask(fn1);
queueMicrotask(fn2);
"#;
    assert_eq!(run_js(src), vec!["Fn1", "Fn2"]);
}

#[test]
fn test_js_queue_microtask_reusing_same_function_reference() {
    let src = r#"
let runs = 0;
const fn = () => { runs++; };
queueMicrotask(fn);
queueMicrotask(fn);
queueMicrotask(() => console.log("Total Runs: " + runs));
"#;
    assert_eq!(run_js(src), vec!["Total Runs: 2"]);
}

#[test]
fn test_js_queue_microtask_in_constructor() {
    let src = r#"
class TaskRunner {
    constructor() {
        queueMicrotask(() => console.log("Constructor Microtask"));
    }
}
new TaskRunner();
console.log("Sync Constructor Done");
"#;
    assert_eq!(
        run_js(src),
        vec!["Sync Constructor Done", "Constructor Microtask"]
    );
}

#[test]
fn test_js_queue_microtask_returns_undefined() {
    let src = r#"
const res = queueMicrotask(() => {});
console.log(res === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_queue_microtask_with_symbol_properties() {
    let src = r#"
const sym = Symbol("id");
const data = { [sym]: "Secret" };
queueMicrotask(() => console.log(data[sym]));
"#;
    assert_eq!(run_js(src), vec!["Secret"]);
}

#[test]
fn test_js_queue_microtask_clearing_local_references() {
    let src = r#"
let temp = { heavyData: [1, 2, 3] };
queueMicrotask(() => {
    temp = null;
    console.log("Cleaned Reference: " + (temp === null));
});
"#;
    assert_eq!(run_js(src), vec!["Cleaned Reference: true"]);
}

#[test]
fn test_js_queue_microtask_execution_before_next_event_loop_turn() {
    let src = r#"
let order = [];
queueMicrotask(() => order.push("Micro1"));
queueMicrotask(() => order.push("Micro2"));
Promise.resolve().then(() => {
    console.log(order.join("->"));
});
"#;
    assert_eq!(run_js(src), vec!["Micro1->Micro2"]);
}
