use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Control flow — if/else, switch, loops, jump statements
// ═══════════════════════════════════════════════════════════

#[test]
fn if_basic() {
    let out = run_csharp(r#"
int x = 5;
if (x > 3) {
    Console.WriteLine("big");
}
"#);
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else() {
    let out = run_csharp(r#"
int x = 2;
if (x > 3) {
    Console.WriteLine("big");
} else {
    Console.WriteLine("small");
}
"#);
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_elseif_chain() {
    let out = run_csharp(r#"
int score = 75;
if (score >= 90) Console.WriteLine("A");
else if (score >= 80) Console.WriteLine("B");
else if (score >= 70) Console.WriteLine("C");
else Console.WriteLine("F");
"#);
    assert_eq!(out, vec!["C"]);
}

#[test]
fn switch_with_break() {
    let out = run_csharp(r#"
int day = 3;
switch (day) {
    case 1: Console.WriteLine("Mon"); break;
    case 2: Console.WriteLine("Tue"); break;
    case 3: Console.WriteLine("Wed"); break;
    default: Console.WriteLine("Other"); break;
}
"#);
    assert_eq!(out, vec!["Wed"]);
}

#[test]
fn switch_default() {
    let out = run_csharp(r#"
int x = 99;
switch (x) {
    case 1: Console.WriteLine("one"); break;
    default: Console.WriteLine("other"); break;
}
"#);
    assert_eq!(out, vec!["other"]);
}

#[test]
fn for_loop() {
    let out = run_csharp(r#"
int sum = 0;
for (int i = 1; i <= 5; i++) {
    sum += i;
}
Console.WriteLine(sum);
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn while_loop() {
    let out = run_csharp(r#"
int i = 0;
while (i < 3) {
    Console.WriteLine(i);
    i++;
}
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn do_while_loop() {
    let out = run_csharp(r#"
int i = 0;
do {
    i++;
} while (i < 3);
Console.WriteLine(i);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
#[ignore]
fn foreach_array() {
    let out = run_csharp(r#"
int[] arr = {10, 20, 30};
int sum = 0;
foreach (var x in arr) {
    sum += x;
}
Console.WriteLine(sum);
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn break_in_loop() {
    let out = run_csharp(r#"
for (int i = 0; i < 100; i++) {
    if (i >= 3) break;
    Console.WriteLine(i);
}
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn continue_in_loop() {
    let out = run_csharp(r#"
int sum = 0;
for (int i = 0; i < 10; i++) {
    if (i % 2 != 0) continue;
    sum += i;
}
Console.WriteLine(sum);
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn nested_loops() {
    let out = run_csharp(r#"
int count = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 4; j++) {
        count++;
    }
}
Console.WriteLine(count);
"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn ternary_expression() {
    let out = run_csharp(r#"
int x = 5;
string result = x > 3 ? "big" : "small";
Console.WriteLine(result);
"#);
    assert_eq!(out, vec!["big"]);
}

#[test]
fn nested_ternary() {
    let out = run_csharp(r#"
int x = 5;
string r = x > 10 ? "big" : x > 3 ? "medium" : "small";
Console.WriteLine(r);
"#);
    assert_eq!(out, vec!["medium"]);
}
