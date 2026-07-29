use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Labeled Statements (`break label`, `continue label`) Control Flow
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_labeled_break_nested_loops() {
    let src = r#"
const res = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (i === 1 && j === 1) break outer;
        res.push(`${i},${j}`);
    }
}
console.log(res.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0,0|0,1|0,2|1,0"]);
}

#[test]
fn test_js_labeled_continue_outer_loop() {
    let src = r#"
const res = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) continue outer;
        res.push(`${i},${j}`);
    }
}
console.log(res.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0,0|1,0|2,0"]);
}

#[test]
fn test_js_labeled_break_block_statement() {
    let src = r#"
const log = [];
log.push("Before");
myBlock: {
    log.push("Inside");
    break myBlock;
    log.push("Unreachable");
}
log.push("After");
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["Before,Inside,After"]);
}

#[test]
fn test_js_labeled_statement_with_switch_break() {
    let src = r#"
const res = [];
outerLoop: for (let i = 1; i <= 2; i++) {
    switch (i) {
        case 1:
            res.push("case1");
            break; // Breaks switch, stays in for loop
        case 2:
            res.push("case2");
            break outerLoop; // Breaks outer for loop!
    }
    res.push("afterSwitch");
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["case1,afterSwitch,case2"]);
}

#[test]
fn test_js_invalid_label_break_throws_syntaxerror() {
    let src = r#"
try {
    eval("break nonExistentLabel;");
} catch (e) {
    console.log("Invalid Break Label SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Break Label SyntaxError"]);
}

#[test]
fn test_js_invalid_label_continue_throws_syntaxerror() {
    let src = r#"
try {
    eval("continue nonExistentLabel;");
} catch (e) {
    console.log("Invalid Continue Label SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Continue Label SyntaxError"]);
}

#[test]
fn test_js_continue_non_loop_label_throws_syntaxerror() {
    let src = r#"
try {
    eval("myBlock: { continue myBlock; }"); // continue can ONLY target loop statements!
} catch (e) {
    console.log("Continue Non-Loop Label SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Continue Non-Loop Label SyntaxError"]);
}

#[test]
fn test_js_duplicate_label_in_nested_scope_throws_syntaxerror() {
    let src = r#"
try {
    eval("label: { label: { break label; } }");
} catch (e) {
    console.log("Duplicate Label SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Duplicate Label SyntaxError"]);
}

#[test]
fn test_js_labeled_do_while_loop() {
    let src = r#"
let i = 0;
const res = [];
loopLabel: do {
    i++;
    if (i === 2) continue loopLabel;
    res.push(i);
} while (i < 3);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_labeled_while_loop_continue_executes_finally_before_iteration() {
    let src = r#"
let i = 0;
const log = [];

outer: while (i < 4) {
    try {
        log.push("body-" + i);
        if (i === 1) {
            i += 1;
            continue outer;
        }
        log.push("work-" + i);
        i++;
    } finally {
        log.push("finally-" + i);
    }
}

console.log(log.join("|"));
"#;

    assert_eq!(
        run_js(src),
        vec!["body-0|work-0|finally-0|body-1|finally-2|work-2|finally-3|work-3|finally-4"]
    );
}

#[test]
fn test_js_labeled_for_in_loop() {
    let src = r#"
const obj = { a: 1, b: 2, c: 3 };
const res = [];
forInLabel: for (const k in obj) {
    if (k === "b") break forInLabel;
    res.push(k);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["a"]);
}

#[test]
fn test_js_labeled_for_of_loop() {
    let src = r#"
const arr = [10, 20, 30];
const res = [];
forOfLabel: for (const val of arr) {
    if (val === 20) continue forOfLabel;
    res.push(val);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["10,30"]);
}

#[test]
fn test_js_multiple_labels_on_single_statement() {
    let src = r#"
let hit = false;
label1: label2: {
    hit = true;
    break label1;
}
console.log(hit);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_labeled_statement_completion_value() {
    let src = r#"
console.log(eval("lbl: { 10; 20; }"));
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_labeled_break_in_finally_block_overrides_return() {
    let src = r#"
function fn() {
    lbl: {
        try {
            return "TryReturn";
        } finally {
            break lbl; // Break label in finally overrides return value!
        }
    }
    return "AfterBlock";
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["AfterBlock"]);
}

#[test]
fn test_js_labeled_continue_in_finally_block_overrides_return() {
    let src = r#"
function fn() {
    const res = [];
    lblLoop: for (let i = 0; i < 2; i++) {
        try {
            return "TryReturn";
        } finally {
            continue lblLoop; // Continue loop in finally overrides return value!
        }
    }
    return "LoopExhausted";
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["LoopExhausted"]);
}

#[test]
fn test_js_labeled_if_statement_block_break() {
    let src = r#"
let executed = false;
ifLabel: if (true) {
    executed = true;
    break ifLabel;
    executed = false;
}
console.log(executed);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_labeled_function_declaration_in_non_strict() {
    let src = r#"
fnLabel: function testFn() { return "LabeledFunc"; }
console.log(testFn());
"#;
    assert_eq!(run_js(src), vec!["LabeledFunc"]);
}

#[test]
fn test_js_labeled_function_declaration_in_strict_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; fnLabel: function testFn() {}");
} catch (e) {
    console.log("Strict Labeled Function SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Labeled Function SyntaxError"]);
}

#[test]
fn test_js_labeled_statement_with_try_catch_unwinding() {
    let src = r#"
const log = [];
lbl: {
    try {
        log.push("Try");
        throw new Error("Err");
    } catch (e) {
        log.push("Catch");
        break lbl;
    } finally {
        log.push("Finally");
    }
    log.push("Unreachable");
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["Try,Catch,Finally"]);
}

#[test]
fn test_js_labeled_statement_scope_shadowing_same_name_different_blocks() {
    let src = r#"
const res = [];
blockA: {
    res.push("A1");
}
blockA: {
    res.push("A2");
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["A1,A2"]);
}

#[test]
fn test_js_labeled_switch_statement_break() {
    let src = r#"
const log = [];
lblSwitch: switch (1) {
    case 1:
        log.push("c1");
        break lblSwitch;
        log.push("unreachable");
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["c1"]);
}

#[test]
fn test_js_multiple_labels_outer_vs_inner_break() {
    let src = r#"
const log = [];
outer: inner: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (i === 1) break outer;
        if (j === 1) break inner;
        log.push(`${i}:${j}`);
    }
}
console.log(log.join("|"));
"#;
    assert_eq!(run_js(src), vec!["0:0|1:0"]);
}

#[test]
fn test_js_labeled_for_of_loop_break() {
    let src = r#"
const arr = [1, 2, 3];
const res = [];
outer: for (const x of arr) {
    if (x === 2) break outer;
    res.push(x);
}
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["1"]);
}


