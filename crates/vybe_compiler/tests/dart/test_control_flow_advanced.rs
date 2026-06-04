use super::helpers::{compile_ok, run_prints};

// ── do-while ─────────────────────────────────────────────────

#[test]
fn do_while_basic() {
    compile_ok("void main() { var i = 0; do { i++; } while (i < 5); }");
}
#[test]
fn do_while_result() {
    let out = run_prints("void main() { var i = 0; do { i++; } while (i < 3); print(i); }");
    assert_eq!(out, ["3"]);
}

#[test]
fn do_while_at_least_once() {
    let out = run_prints("void main() { var i = 10; do { print(i); i++; } while (i < 5); }");
    assert_eq!(out, ["10"]);
}

// ── break ────────────────────────────────────────────────────

#[test]
fn break_in_for() {
    compile_ok("void main() { for (var i = 0; i < 10; i++) { if (i == 5) break; } }");
}
#[test]
fn break_in_while() {
    compile_ok("void main() { var i = 0; while (true) { if (i >= 5) break; i++; } }");
}
#[test]
fn break_in_for_in() {
    compile_ok("void main() { for (var x in [1,2,3,4,5]) { if (x == 3) break; } }");
}

#[test]
fn break_result() {
    let out = run_prints(
        r#"
void main() {
  var sum = 0;
  for (var i = 1; i <= 10; i++) {
    if (i > 5) break;
    sum += i;
  }
  print(sum);
}
"#,
    );
    assert_eq!(out, ["15"]);
}

// ── continue ─────────────────────────────────────────────────

#[test]
fn continue_in_for() {
    compile_ok(
        "void main() { for (var i = 0; i < 10; i++) { if (i % 2 == 0) continue; print(i); } }",
    );
}
#[test]
fn continue_in_while() {
    compile_ok(
        "void main() { var i = 0; while (i < 10) { i++; if (i % 2 == 0) continue; print(i); } }",
    );
}

#[test]
fn continue_result() {
    let out = run_prints(
        r#"
void main() {
  var odds = <int>[];
  for (var i = 1; i <= 6; i++) {
    if (i % 2 == 0) continue;
    odds.add(i);
  }
  print(odds.length);
}
"#,
    );
    assert_eq!(out, ["3"]);
}

// ── Nested loops ─────────────────────────────────────────────

#[test]
fn nested_for() {
    compile_ok(
        "void main() { for (var i = 0; i < 3; i++) { for (var j = 0; j < 3; j++) { print('$i,$j'); } } }",
    );
}

#[test]
fn nested_result() {
    let out = run_prints(
        r#"
void main() {
  var count = 0;
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      count++;
    }
  }
  print(count);
}
"#,
    );
    assert_eq!(out, ["9"]);
}

#[test]
fn break_inner_loop() {
    let out = run_prints(
        r#"
void main() {
  var count = 0;
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      if (j == 1) break;
      count++;
    }
  }
  print(count);
}
"#,
    );
    assert_eq!(out, ["3"]);
}

// ── switch with complex cases ────────────────────────────────

#[test]
fn switch_multiple_values() {
    compile_ok(
        r#"
void main() {
  var x = 'b';
  switch (x) {
    case 'a':
    case 'b':
      print('vowel-ish');
      break;
    default:
      print('other');
  }
}
"#,
    );
}

#[test]
fn switch_string() {
    let out = run_prints(
        r#"
void main() {
  var day = 'Mon';
  switch (day) {
    case 'Mon': print('Monday'); break;
    case 'Tue': print('Tuesday'); break;
    default: print('Other');
  }
}
"#,
    );
    assert_eq!(out, ["Monday"]);
}

#[test]
fn switch_int() {
    let out = run_prints(
        r#"
void main() {
  var n = 2;
  switch (n) {
    case 1: print('one'); break;
    case 2: print('two'); break;
    case 3: print('three'); break;
    default: print('many');
  }
}
"#,
    );
    assert_eq!(out, ["two"]);
}

#[test]
fn switch_default_only() {
    let out = run_prints(
        r#"
void main() {
  switch (99) {
    default: print('caught');
  }
}
"#,
    );
    assert_eq!(out, ["caught"]);
}

#[test]
fn switch_no_match() {
    let out = run_prints(
        r#"
void main() {
  var x = 5;
  switch (x) {
    case 1: print('one'); break;
    case 2: print('two'); break;
    default: print('other');
  }
}
"#,
    );
    assert_eq!(out, ["other"]);
}

// ── if / else-if chains ──────────────────────────────────────

#[test]
fn else_if_chain() {
    let out = run_prints(
        r#"
void main() {
  var score = 75;
  if (score >= 90) {
    print('A');
  } else if (score >= 80) {
    print('B');
  } else if (score >= 70) {
    print('C');
  } else {
    print('F');
  }
}
"#,
    );
    assert_eq!(out, ["C"]);
}

#[test]
fn nested_if() {
    let out = run_prints(
        r#"
void main() {
  var x = 5;
  var y = 10;
  if (x > 0) {
    if (y > 0) {
      print('both positive');
    }
  }
}
"#,
    );
    assert_eq!(out, ["both positive"]);
}

// ── while loop variants ──────────────────────────────────────

#[test]
fn while_accumulate() {
    let out = run_prints(
        r#"
void main() {
  var i = 1;
  var product = 1;
  while (i <= 5) {
    product *= i;
    i++;
  }
  print(product);
}
"#,
    );
    assert_eq!(out, ["120"]);
}

#[test]
fn while_string_build() {
    let out = run_prints(
        r#"
void main() {
  var s = '';
  var i = 0;
  while (i < 3) {
    s += 'a';
    i++;
  }
  print(s);
}
"#,
    );
    assert_eq!(out, ["aaa"]);
}

// ── for-in variants ──────────────────────────────────────────

#[test]
fn for_in_map_keys() {
    compile_ok("void main() { var m = {'a': 1, 'b': 2}; for (var k in m.keys) { print(k); } }");
}

#[test]
fn for_in_set() {
    compile_ok("void main() { var s = {1, 2, 3}; for (var x in s) { print(x); } }");
}

#[test]
fn for_in_collect() {
    let out = run_prints(
        r#"
void main() {
  var sum = 0;
  for (var x in [1, 2, 3, 4, 5]) {
    sum += x;
  }
  print(sum);
}
"#,
    );
    assert_eq!(out, ["15"]);
}

// ── Chained conditions ───────────────────────────────────────

#[test]
fn and_condition() {
    let out = run_prints("void main() { var a = 5; var b = 10; print(a > 0 && b > 0); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn or_condition() {
    let out = run_prints("void main() { var a = -1; var b = 10; print(a > 0 || b > 0); }");
    assert_eq!(out, ["true"]);
}

#[test]
fn not_condition() {
    let out = run_prints("void main() { var flag = false; if (!flag) { print('yes'); } }");
    assert_eq!(out, ["yes"]);
}

// ── Ranges and counting ──────────────────────────────────────

#[test]
fn count_up() {
    let out = run_prints(
        r#"
void main() {
  var result = <int>[];
  for (var i = 1; i <= 5; i++) { result.add(i); }
  print(result.length);
}
"#,
    );
    assert_eq!(out, ["5"]);
}

#[test]
fn count_down() {
    let out = run_prints(
        r#"
void main() {
  var result = <int>[];
  for (var i = 5; i >= 1; i--) { result.add(i); }
  print(result.first);
}
"#,
    );
    assert_eq!(out, ["5"]);
}
