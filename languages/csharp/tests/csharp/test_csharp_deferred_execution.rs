//! LINQ deferred execution, materialization, and side-effect ordering.
use super::helpers::run_csharp;

#[test]
fn where_query_not_executed_until_enumerated() {
    assert_eq!(
        run_csharp(
            r#"int count=0;
var q=new[]{1,2,3}.Where(n=>{count++;return n>1;});
Console.WriteLine(count);
var list=q.ToList();
Console.WriteLine(count);"#
        ),
        &["0", "3"]
    );
}

#[test]
fn select_query_executes_on_each_enumeration() {
    assert_eq!(
        run_csharp(
            r#"int calls=0;
var q=new[]{1,2,3}.Select(n=>{calls++;return n*2;});
var r1=q.ToList(); var r2=q.ToList();
Console.WriteLine(calls);"#
        ),
        &["6"]
    );
}

#[test]
fn to_list_materialises_and_snapshot_captures_source() {
    assert_eq!(
        run_csharp(
            r#"var source=new System.Collections.Generic.List<int>{1,2,3};
var snapshot=source.ToList();
source.Add(4);
Console.WriteLine(snapshot.Count);"#
        ),
        &["3"]
    );
}

#[test]
fn any_short_circuits_after_first_match() {
    assert_eq!(
        run_csharp(
            r#"int count=0;
bool found=new[]{1,2,3,4,5}.Any(n=>{count++;return n==3;});
Console.WriteLine(found); Console.WriteLine(count);"#
        ),
        &["True", "3"]
    );
}

#[test]
fn first_throws_if_sequence_is_empty() {
    assert_eq!(
        run_csharp(
            r#"string r="";
try{System.Array.Empty<int>().First();}
catch(System.InvalidOperationException){r="empty";}
Console.WriteLine(r);"#
        ),
        &["empty"]
    );
}

#[test]
fn count_vs_to_list_count_return_same_number() {
    assert_eq!(
        run_csharp(
            r#"var q=new[]{1,2,3,4}.Where(x=>x%2==0);
Console.WriteLine(q.Count());
Console.WriteLine(q.ToList().Count);"#
        ),
        &["2", "2"]
    );
}
