//! Nested loops, `break`/`continue`, and `switch` control transfer with
//! observable side effects on outer state.
use super::helpers::run_csharp;

#[test]
fn break_inside_inner_loop_does_not_stop_outer_loop() {
    assert_eq!(
        run_csharp(
            r#"
int total = 0;
for (int row = 0; row < 2; row++) {
    for (int col = 0; col < 4; col++) {
        if (col == 2) break;
        total += 1;
    }
}
Console.WriteLine(total);
"#
        ),
        &["4"]
    );
}

#[test]
fn continue_inside_inner_loop_skips_remaining_body_but_not_outer() {
    assert_eq!(
        run_csharp(
            r#"
int sum = 0;
for (int outer = 0; outer < 2; outer++) {
    for (int inner = 0; inner < 3; inner++) {
        if (inner == 1) continue;
        sum += inner;
    }
}
Console.WriteLine(sum);
"#
        ),
        &["4"]
    );
}

#[test]
fn foreach_break_exits_after_first_matching_element() {
    assert_eq!(
        run_csharp(
            r#"
int hits = 0;
foreach (var value in new[] { 2, 4, 6, 8 }) {
    if (value == 6) break;
    hits++;
}
Console.WriteLine(hits);
"#
        ),
        &["2"]
    );
}

#[test]
fn foreach_continue_skips_even_numbers_only() {
    assert_eq!(
        run_csharp(
            r#"
int sum = 0;
foreach (var value in new[] { 1, 2, 3, 4, 5 }) {
    if (value % 2 == 0) continue;
    sum += value;
}
Console.WriteLine(sum);
"#
        ),
        &["9"]
    );
}

#[test]
fn while_loop_with_continue_reaches_next_condition_check() {
    assert_eq!(
        run_csharp(
            r#"
int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n == 3) continue;
    sum += n;
}
Console.WriteLine(sum);
"#
        ),
        &["12"]
    );
}

#[test]
fn do_while_executes_body_before_first_condition_test() {
    assert_eq!(
        run_csharp(
            r#"
int count = 0;
do {
    count++;
} while (count < 1);
Console.WriteLine(count);
"#
        ),
        &["1"]
    );
}

#[test]
fn switch_break_prevents_fallthrough_into_next_case() {
    assert_eq!(
        run_csharp(
            r#"
int code = 2;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    case 3: label = "three"; break;
}
Console.WriteLine(label);
"#
        ),
        &["two"]
    );
}

#[test]
fn switch_goto_case_runs_second_case_after_first_match() {
    assert_eq!(
        run_csharp(
            r#"
int code = 1;
string trace = "";
switch (code) {
    case 1:
        trace += "A";
        goto case 2;
    case 2:
        trace += "B";
        break;
}
Console.WriteLine(trace);
"#
        ),
        &["AB"]
    );
}

#[test]
fn nested_switch_break_exits_only_inner_switch() {
    assert_eq!(
        run_csharp(
            r#"
string report = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0:
                    report += "inner;";
                    break;
            }
            report += "after-inner;";
            break;
        case 1:
            report += "tail;";
            break;
    }
}
Console.WriteLine(report);
"#
        ),
        &["inner;after-inner;tail;"]
    );
}

#[test]
fn for_loop_initializer_scope_does_not_leak_after_loop() {
    assert_eq!(
        run_csharp(
            r#"
int total = 0;
for (int i = 0; i < 3; i++) total += i;
Console.WriteLine(total);
"#
        ),
        &["3"]
    );
}

#[test]
fn foreach_iteration_variable_is_fresh_each_iteration() {
    assert_eq!(
        run_csharp(
            r#"
int last = -1;
foreach (var value in new[] { 1, 2, 3 }) {
    last = value;
}
Console.WriteLine(last);
"#
        ),
        &["3"]
    );
}

#[test]
fn switch_break_inside_loop_allows_subsequent_iterations() {
    assert_eq!(
        run_csharp(
            r#"
int sum = 0;
for (int i = 0; i < 4; i++) {
    switch (i) {
        case 1:
        case 2:
            sum += 10;
            break;
        default:
            sum += 1;
            break;
    }
}
Console.WriteLine(sum);
"#
        ),
        &["22"]
    );
}
