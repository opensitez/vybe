//! Generic methods on non-generic classes, generic delegates, and type inference.
use super::helpers::run_csharp;

#[test]
fn generic_method_on_non_generic_class_infers_type() {
    assert_eq!(
        run_csharp(
            r#"class Utils{public static T First<T>(T[] arr)=>arr[0];}
Console.WriteLine(Utils.First(new[]{10,20,30}));
Console.WriteLine(Utils.First(new[]{"a","b"}));"#
        ),
        &["10", "a"]
    );
}

#[test]
fn generic_method_with_explicit_type_argument() {
    assert_eq!(
        run_csharp(
            r#"T Box<T>(T v)=>v;
Console.WriteLine(Box<int>(5));"#
        ),
        &["5"]
    );
}

#[test]
fn generic_method_filters_sequence_by_type() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<T> FilterType<T>(object[] items){
    foreach(var i in items) if(i is T t) yield return t;
}
var items=new object[]{1,"a",2,"b",3};
int count=0;
foreach(var s in FilterType<string>(items)) count++;
Console.WriteLine(count);"#
        ),
        &["2"]
    );
}

#[test]
fn generic_method_swap_exchanges_two_values_via_ref() {
    assert_eq!(
        run_csharp(
            r#"void Swap<T>(ref T a,ref T b){T tmp=a;a=b;b=tmp;}
int x=1,y=2; Swap(ref x,ref y);
Console.WriteLine(x); Console.WriteLine(y);"#
        ),
        &["2", "1"]
    );
}

#[test]
fn generic_action_parameterised_with_type_argument() {
    assert_eq!(
        run_csharp(
            r#"void ForEach<T>(T[] items,System.Action<T> action){
    foreach(var i in items) action(i);
}
ForEach(new[]{1,2,3},n=>Console.WriteLine(n));"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn generic_func_composes_two_functions() {
    assert_eq!(
        run_csharp(
            r#"System.Func<A,C> Compose<A,B,C>(System.Func<A,B> f,System.Func<B,C> g)=>x=>g(f(x));
var fn=Compose((int x)=>x*2,(int y)=>y+1);
Console.WriteLine(fn(5));"#
        ),
        &["11"]
    );
}

#[test]
fn generic_method_returns_default_for_empty_sequence() {
    assert_eq!(
        run_csharp(
            r#"T FirstOrDefault<T>(T[] arr)=>arr.Length>0?arr[0]:default;
Console.WriteLine(FirstOrDefault(new int[]{}));
Console.WriteLine(FirstOrDefault(new[]{9}));"#
        ),
        &["0", "9"]
    );
}
