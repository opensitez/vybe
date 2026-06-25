//! Advanced `yield`: `yield break`, `yield` in `try/finally`, nested iterators.
use super::helpers::run_csharp;

#[test]
fn yield_break_stops_iteration_early() {
    assert_eq!(
        run_csharp(r#"System.Collections.Generic.IEnumerable<int> Take(int[] a,int max){
    int count=0;
    foreach(var n in a){
        if(count>=max) yield break;
        yield return n;
        count++;
    }
}
Console.WriteLine(string.Join(",",Take(new[]{1,2,3,4,5},3)));"#),
        &["1,2,3"]
    );
}

#[test]
fn yield_in_try_finally_disposes_after_iteration() {
    assert_eq!(
        run_csharp(r#"bool cleaned=false;
System.Collections.Generic.IEnumerable<int> Gen(){
    try{ yield return 1; yield return 2; }
    finally{ cleaned=true; }
}
foreach(var _ in Gen()){}
Console.WriteLine(cleaned);"#),
        &["True"]
    );
}

#[test]
fn nested_iterators_produce_flat_result_when_chained() {
    assert_eq!(
        run_csharp(r#"System.Collections.Generic.IEnumerable<int> Doubles(int n){
    yield return n; yield return n*2;
}
var result=new[]{1,2,3}.SelectMany(Doubles);
Console.WriteLine(string.Join(",",result));"#),
        &["1,2,2,4,3,6"]
    );
}

#[test]
fn lazy_generator_only_computes_needed_values() {
    assert_eq!(
        run_csharp(r#"int calls=0;
System.Collections.Generic.IEnumerable<int> Expensive(){
    for(int i=0;;i++){calls++;yield return i;}
}
var first3=Expensive().Take(3).ToList();
Console.WriteLine(calls); Console.WriteLine(first3[2]);"#),
        &["3", "2"]
    );
}

#[test]
fn yield_return_with_complex_state_machine() {
    assert_eq!(
        run_csharp(r#"System.Collections.Generic.IEnumerable<string> Words(string s){
    var parts=s.Split(' ');
    foreach(var p in parts) if(p.Length>0) yield return p;
}
Console.WriteLine(string.Join("|",Words("hello  world  foo")));"#),
        &["hello|world|foo"]
    );
}
