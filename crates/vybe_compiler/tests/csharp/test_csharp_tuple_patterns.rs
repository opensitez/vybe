//! Tuple patterns in switch, deconstruction assignment, and nested patterns.
use super::helpers::run_csharp;

#[test]
fn tuple_pattern_in_switch_expression_dispatches_on_pair() {
    assert_eq!(
        run_csharp(
            r#"string Classify(int x,int y)=>(x,y) switch{
    (0,0)=>"origin",
    (>0,0)=>"pos-x",
    (0,>0)=>"pos-y",
    _=>"other"};
Console.WriteLine(Classify(0,0));
Console.WriteLine(Classify(3,0));
Console.WriteLine(Classify(0,5));
Console.WriteLine(Classify(1,1));"#
        ),
        &["origin", "pos-x", "pos-y", "other"]
    );
}

#[test]
fn tuple_deconstruction_assigns_to_multiple_locals() {
    assert_eq!(
        run_csharp(
            r#"(string name,int age)=("Alice",30);
Console.WriteLine(name); Console.WriteLine(age);"#
        ),
        &["Alice", "30"]
    );
}

#[test]
fn tuple_deconstruction_with_discard() {
    assert_eq!(
        run_csharp(
            r#"(string name,_,int score)=("Bob",99,"skip",55) switch{
    var t=>(t.Item1,t.Item2,t.Item4)};
Console.WriteLine(name); Console.WriteLine(score);"#
        ),
        &["Bob", "55"]
    );
}

#[test]
fn nested_tuple_pattern_matches_inner_value() {
    assert_eq!(
        run_csharp(
            r#"var data=((1,2),(3,4));
var((a,b),(c,d))=data;
Console.WriteLine(a+b+c+d);"#
        ),
        &["10"]
    );
}

#[test]
fn tuple_swap_without_temp_variable() {
    assert_eq!(
        run_csharp(
            r#"int x=1,y=2;
(x,y)=(y,x);
Console.WriteLine(x); Console.WriteLine(y);"#
        ),
        &["2", "1"]
    );
}
