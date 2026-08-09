/// Label statements, break/continue with labels, nested loops
use super::helpers::run_js;

#[test]
fn labeled_break_exits_outer_loop() {
    assert_eq!(
        run_js(
            r#"
let result = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) break outer;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
"#
        ),
        vec!["0,0"]
    );
}

#[test]
fn labeled_continue_skips_outer_iteration() {
    assert_eq!(
        run_js(
            r#"
let result = [];
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) continue outer;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
"#
        ),
        vec!["0,0|1,0|2,0"]
    );
}

#[test]
fn labeled_continue_from_inner_try_catch() {
    assert_eq!(
        run_js(
            r#"
let log = [];
outer: for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) {
            throw "skip";
        }
        log.push("ok" + i);
    } catch {
        log.push("catch" + i);
        continue outer;
    }
    log.push("done" + i);
}
console.log(log.join("|"));
"#
        ),
        vec!["ok0|done0|catch1|ok2|done2"]
    );
}

#[test]
fn unlabeled_break_exits_inner_only() {
    assert_eq!(
        run_js(
            r#"
let result = [];
for (let i = 0; i < 2; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) break;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
"#
        ),
        vec!["0,0|1,0"]
    );
}

#[test]
fn break_exits_switch_not_loop() {
    assert_eq!(
        run_js(
            r#"
let result = [];
for (let i = 0; i < 3; i++) {
    switch (i) {
        case 1: break; // exits switch, not loop
    }
    result.push(i);
}
console.log(result.join(","));
"#
        ),
        vec!["0,1,2"]
    );
}

#[test]
fn labeled_break_exits_labeled_block() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
block: {
    x = 1;
    break block;
    x = 2; // never reached
}
console.log(x);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn continue_in_while_loop() {
    assert_eq!(
        run_js(
            r#"
let i = 0, sum = 0;
while (i < 10) {
    i++;
    if (i % 2 === 0) continue;
    sum += i;
}
console.log(sum); // 1+3+5+7+9 = 25
"#
        ),
        vec!["25"]
    );
}

#[test]
fn break_in_do_while() {
    assert_eq!(
        run_js(
            r#"
let i = 0;
do {
    if (i === 3) break;
    i++;
} while (true);
console.log(i);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn labeled_for_of() {
    assert_eq!(
        run_js(
            r#"
let found = null;
outer: for (const arr of [[1,2],[3,4],[5,6]]) {
    for (const x of arr) {
        if (x === 4) { found = x; break outer; }
    }
}
console.log(found);
"#
        ),
        vec!["4"]
    );
}

#[test]
fn triple_nested_labeled() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
a: for (let i = 0; i < 3; i++) {
    b: for (let j = 0; j < 3; j++) {
        c: for (let k = 0; k < 3; k++) {
            if (k === 1) continue b;
            count++;
        }
    }
}
console.log(count); // each i,j pair contributes 1 (k=0), skips rest
"#
        ),
        vec!["9"]
    );
}

#[test]
fn for_in_with_break() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
let first = null;
for (const key in obj) {
    first = key;
    break;
}
console.log(first);
"#
        ),
        vec!["a"]
    );
}

#[test]
fn label_does_not_create_scope() {
    assert_eq!(
        run_js(
            r#"
let x = 0;
myLabel: {
    let y = 10; // block scope, not label scope
    x = y;
}
console.log(x);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn labeled_break_nested_blocks() {
    let src = r#"
let res = "none";
outer: {
    inner: {
        res = "inner";
        break outer;
        res = "after";
    }
    res = "outer";
}
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["inner"]);
}
