//! Lambda expression forms: statement bodies, expression bodies, captures, generics.
use super::helpers::run_csharp;

#[test]
fn expression_lambda_returns_computed_result() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int> f = x => x*x;
Console.WriteLine(f(5));"#
        ),
        &["25"]
    );
}

#[test]
fn statement_lambda_body_executes_multiple_lines() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int> fact = null;
fact = n => { if(n<=1) return 1; return n*fact(n-1); };
Console.WriteLine(fact(5));"#
        ),
        &["120"]
    );
}

#[test]
fn lambda_passed_as_argument_to_higher_order_method() {
    assert_eq!(
        run_csharp(
            r#"int Apply(System.Func<int,int,int> op, int a, int b) => op(a,b);
Console.WriteLine(Apply((a,b) => a+b, 3, 4));"#
        ),
        &["7"]
    );
}

#[test]
fn lambda_with_no_parameters_using_empty_parens() {
    assert_eq!(
        run_csharp(
            r#"System.Func<string> greeting = () => "hello";
Console.WriteLine(greeting());"#
        ),
        &["hello"]
    );
}

#[test]
fn lambda_captures_and_modifies_outer_list() {
    assert_eq!(
        run_csharp(
            r#"var results = new System.Collections.Generic.List<int>();
var nums = new[]{1,2,3,4};
System.Array.ForEach(nums, n => { if(n%2==0) results.Add(n); });
Console.WriteLine(results.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn lambda_returning_lambda_builds_curried_function() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,System.Func<int,int>> add = a => b => a+b;
var add5 = add(5);
Console.WriteLine(add5(3));"#
        ),
        &["8"]
    );
}

#[test]
fn linq_where_takes_lambda_as_predicate() {
    assert_eq!(
        run_csharp(
            r#"var evens = new[]{1,2,3,4,5,6}.Where(n => n%2==0);
Console.WriteLine(evens.Count());"#
        ),
        &["3"]
    );
}

#[test]
fn lambda_implicitly_typed_with_var_in_local_variable() {
    assert_eq!(
        run_csharp(
            r#"var f = (int x) => x + 1;
Console.WriteLine(f(9));"#
        ),
        &["10"]
    );
}
