//! Closure variable capture: outer locals, loop variables, mutation, static closures.
use super::helpers::run_csharp;

#[test]
fn closure_captures_enclosing_local_by_reference() {
    assert_eq!(
        run_csharp(
            r#"int x = 1;
System.Action inc = () => x++;
inc(); inc();
Console.WriteLine(x);"#
        ),
        &["3"]
    );
}

#[test]
fn closure_sees_mutation_of_captured_variable() {
    assert_eq!(
        run_csharp(
            r#"int x = 0;
System.Func<int> read = () => x;
x = 99;
Console.WriteLine(read());"#
        ),
        &["99"]
    );
}

#[test]
fn foreach_closure_captures_correct_loop_variable_with_local_copy() {
    assert_eq!(
        run_csharp(
            r#"var actions = new System.Collections.Generic.List<System.Func<int>>();
foreach(var v in new[]{10,20,30}) {
    var copy = v;
    actions.Add(() => copy);
}
foreach(var a in actions) Console.WriteLine(a());"#
        ),
        &["10", "20", "30"]
    );
}

#[test]
fn for_loop_capture_of_loop_variable_with_local_copy() {
    assert_eq!(
        run_csharp(
            r#"var actions = new System.Collections.Generic.List<System.Func<int>>();
for(int i=0; i<3; i++) {
    var copy = i;
    actions.Add(() => copy);
}
foreach(var a in actions) Console.WriteLine(a());"#
        ),
        &["0", "1", "2"]
    );
}

#[test]
fn two_closures_share_same_captured_variable() {
    assert_eq!(
        run_csharp(
            r#"int shared = 0;
System.Action add = () => shared++;
System.Func<int> read = () => shared;
add(); add();
Console.WriteLine(read());"#
        ),
        &["2"]
    );
}

#[test]
fn nested_closure_captures_from_outer_scope() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,System.Func<int>> makeAdder = x => () => x + 1;
var add1 = makeAdder(5);
Console.WriteLine(add1());"#
        ),
        &["6"]
    );
}
