//! Functional patterns: pipeline chaining, immutability idioms, higher-order functions.
use super::helpers::run_csharp;

#[test]
fn method_chaining_builds_fluent_pipeline() {
    assert_eq!(
        run_csharp(r#"var result=new[]{5,3,8,1,4}
    .Where(x=>x>2)
    .OrderBy(x=>x)
    .Select(x=>x*10)
    .First();
Console.WriteLine(result);"#),
        &["30"]
    );
}

#[test]
fn reduce_via_aggregate_applies_binary_operation() {
    assert_eq!(
        run_csharp(r#"var product=new[]{1,2,3,4,5}.Aggregate((acc,x)=>acc*x);
Console.WriteLine(product);"#),
        &["120"]
    );
}

#[test]
fn map_then_filter_then_reduce_pipeline() {
    assert_eq!(
        run_csharp(r#"var result=new[]{1,2,3,4,5}
    .Select(x=>x*x)
    .Where(x=>x>5)
    .Sum();
Console.WriteLine(result);"#),
        &["50"]
    );
}

#[test]
fn function_composition_applies_in_sequence() {
    assert_eq!(
        run_csharp(r#"System.Func<int,int> triple=x=>x*3;
System.Func<int,int> addOne=x=>x+1;
var composed=new[]{1,2,3}.Select(triple).Select(addOne);
foreach(var n in composed) Console.WriteLine(n);"#),
        &["4", "7", "10"]
    );
}

#[test]
fn memoize_caches_expensive_computation() {
    assert_eq!(
        run_csharp(r#"var cache=new System.Collections.Generic.Dictionary<int,int>();
System.Func<int,int> fib=null;
fib=n=>{
    if(n<=1) return n;
    if(cache.TryGetValue(n,out int v)) return v;
    var r=fib(n-1)+fib(n-2);
    cache[n]=r;
    return r;
};
Console.WriteLine(fib(10));"#),
        &["55"]
    );
}

#[test]
fn partial_application_creates_specialized_function() {
    assert_eq!(
        run_csharp(r#"System.Func<int,System.Func<int,int>> add=a=>b=>a+b;
var add10=add(10);
Console.WriteLine(add10(5));
Console.WriteLine(add10(20));"#),
        &["15", "30"]
    );
}

#[test]
fn unfold_pattern_generates_fibonacci_via_iteration() {
    assert_eq!(
        run_csharp(r#"System.Collections.Generic.IEnumerable<int> Fibs(){
    int a=0,b=1;
    while(true){yield return a; (a,b)=(b,a+b);}
}
var first8=Fibs().Take(8).ToArray();
Console.WriteLine(first8[7]);"#),
        &["13"]
    );
}
