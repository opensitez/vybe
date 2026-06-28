//! Advanced delegate patterns: currying, memoization, event-like pipelines.
use super::helpers::run_csharp;

#[test]
fn delegate_stored_in_dictionary_and_dispatched_by_key() {
    assert_eq!(
        run_csharp(
            r#"var ops=new System.Collections.Generic.Dictionary<string,System.Func<int,int,int>>{
    {"add",(a,b)=>a+b},
    {"mul",(a,b)=>a*b}
};
Console.WriteLine(ops["add"](3,4));
Console.WriteLine(ops["mul"](3,4));"#
        ),
        &["7", "12"]
    );
}

#[test]
fn chained_func_composition() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,int> double_it=x=>x*2;
System.Func<int,int> add_three=x=>x+3;
System.Func<int,int> combined=x=>add_three(double_it(x));
Console.WriteLine(combined(5));"#
        ),
        &["13"]
    );
}

#[test]
fn predicate_combined_with_and() {
    assert_eq!(
        run_csharp(
            r#"System.Predicate<int> positive=x=>x>0;
System.Predicate<int> even=x=>x%2==0;
System.Predicate<int> both=x=>positive(x)&&even(x);
Console.WriteLine(both(4)); Console.WriteLine(both(-2)); Console.WriteLine(both(3));"#
        ),
        &["True", "False", "False"]
    );
}

#[test]
fn lambda_closed_over_mutable_list_builds_result() {
    assert_eq!(
        run_csharp(
            r#"var log=new System.Collections.Generic.List<string>();
System.Action<string> record=msg=>log.Add(msg);
record("a"); record("b"); record("c");
Console.WriteLine(string.Join(",",log));"#
        ),
        &["a,b,c"]
    );
}

#[test]
fn func_returns_func_partial_application() {
    assert_eq!(
        run_csharp(
            r#"System.Func<int,System.Func<int,int>> multiply=factor=>n=>n*factor;
var triple=multiply(3);
Console.WriteLine(triple(7));"#
        ),
        &["21"]
    );
}
