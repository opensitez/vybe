//! Closures over loop variables: `for` shares one slot across deferred calls,
//! `foreach` iteration variable is per-iteration, manual copy breaks sharing.
use super::helpers::run_csharp;

#[test]
fn for_loop_lambda_captures_shared_counter_showing_last_value_at_invoke_time() {
    assert_eq!(
        run_csharp(
            r#"
using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    actions.Add(() => i);
}
foreach (var run in actions) Console.WriteLine(run());
"#
        ),
        &["3", "3", "3"]
    );
}

#[test]
fn foreach_iteration_lambda_sees_each_elements_value_not_final_index() {
    assert_eq!(
        run_csharp(
            r#"
using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
foreach (var value in new[] { 10, 20, 30 }) {
    actions.Add(() => value);
}
foreach (var run in actions) Console.WriteLine(run());
"#
        ),
        &["10", "20", "30"]
    );
}

#[test]
fn explicit_loop_copy_variable_gives_distinct_closure_per_iteration() {
    assert_eq!(
        run_csharp(
            r#"
using System;
using System.Collections.Generic;
var actions = new List<Func<int>>();
for (int i = 0; i < 3; i++) {
    int copy = i;
    actions.Add(() => copy);
}
foreach (var run in actions) Console.WriteLine(run());
"#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn lambda_mutating_captured_local_is_visible_to_later_invocations() {
    assert_eq!(
        run_csharp(
            r#"
using System;
int tally = 0;
Action bump = () => { tally++; };
bump();
bump();
Console.WriteLine(tally);
"#
        ),
        &["2"]
    );
}

#[test]
fn lambda_passed_to_method_receives_current_binding_not_snapshot_at_creation() {
    assert_eq!(
        run_csharp(
            r#"
using System;
int total = 1;
Action add = () => total += 4;
void Apply(Action work) { work(); }
Apply(add);
Console.WriteLine(total);
"#
        ),
        &["5"]
    );
}
