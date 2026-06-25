//! `Action`, `Func`, `Predicate` and delegate invocation semantics.
use super::helpers::run_csharp;

#[test]
fn action_delegate_calls_void_method_with_no_args() {
    assert_eq!(
        run_csharp(
            r#"System.Action greet = () => Console.WriteLine("hi");
greet();"#
        ),
        &["hi"]
    );
}

#[test]
fn action_t_carries_a_typed_argument() {
    assert_eq!(
        run_csharp(
            r#"System.Action<int> print = n => Console.WriteLine(n * 2);
print(5);"#
        ),
        &["10"]
    );
}

#[test]
fn func_t_result_returns_computed_value() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int> square = x => x * x;
Console.WriteLine(square(4));"#
        ),
        &["16"]
    );
}

#[test]
fn func_t1_t2_result_takes_two_args() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int,int> add = (a,b) => a+b;
Console.WriteLine(add(3,4));"#
        ),
        &["7"]
    );
}

#[test]
fn predicate_t_tests_condition_on_value() {
    assert_eq!(
        run_csharp(
            r#"System.Predicate<string> isLong = s => s.Length > 4;
Console.WriteLine(isLong("hello"));
Console.WriteLine(isLong("hi"));"#
        ),
        &["True", "False"]
    );
}

#[test]
fn method_group_assigned_to_func_without_lambda_wrapper() {
    assert_eq!(
        run_csharp(
            r#"System.Func<string,int> len = s => s.Length;
Console.WriteLine(len("test"));"#
        ),
        &["4"]
    );
}

#[test]
fn multicast_delegate_invokes_both_handlers_in_order() {
    assert_eq!(
        run_csharp(
            r#"System.Action log = () => Console.WriteLine("a");
log += () => Console.WriteLine("b");
log();"#
        ),
        &["a", "b"]
    );
}

#[test]
fn removing_handler_from_multicast_leaves_remaining() {
    assert_eq!(
        run_csharp(
            r#"int count = 0;
System.Action a = () => count++;
System.Action b = () => count++;
System.Action multi = a;
multi += b;
multi -= a;
multi();
Console.WriteLine(count);"#
        ),
        &["1"]
    );
}

#[test]
fn delegate_null_check_before_invoke_prevents_null_reference() {
    assert_eq!(
        run_csharp(
            r#"System.Action handler = null;
handler?.Invoke();
Console.WriteLine("safe");"#
        ),
        &["safe"]
    );
}

#[test]
fn anonymous_method_syntax_works_as_delegate_body() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int> triple = delegate(int n) { return n * 3; };
Console.WriteLine(triple(3));"#
        ),
        &["9"]
    );
}

#[test]
fn func_stored_in_variable_and_passed_to_method() {
    assert_eq!(
        run_csharp(
            r#"int Apply(System.Func<int,int> f, int v) => f(v);
Console.WriteLine(Apply(x => x + 1, 9));"#
        ),
        &["10"]
    );
}
