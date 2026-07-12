//! Advanced pattern-matching switch: guards, relational combos, nested, exhaustive.
use super::helpers::run_csharp;

#[test]
fn switch_expression_with_when_guard() {
    assert_eq!(
        run_csharp(
            r#"string Classify(int n)=>n switch{
    int x when x<0=>"negative",
    0=>"zero",
    int x when x%2==0=>"even",
    _=>"odd"};
Console.WriteLine(Classify(-5));
Console.WriteLine(Classify(0));
Console.WriteLine(Classify(4));
Console.WriteLine(Classify(7));"#
        ),
        &["negative", "zero", "even", "odd"]
    );
}

#[test]
fn nested_tuple_pattern_matches_pair_of_conditions() {
    assert_eq!(
        run_csharp(
            r#"string Combo(bool a,bool b)=>(a,b) switch{
    (true,true)=>"both",
    (true,false)=>"left",
    (false,true)=>"right",
    _=>"none"};
Console.WriteLine(Combo(true,false));"#
        ),
        &["left"]
    );
}

#[test]
fn list_pattern_matches_exact_element_count() {
    assert_eq!(
        run_csharp(
            r#"string Check(int[] a)=>a switch{
    []=>"empty",
    [_]=>"single",
    [_,_]=>"pair",
    _=>"many"};
Console.WriteLine(Check(new int[]{}));
Console.WriteLine(Check(new[]{1}));
Console.WriteLine(Check(new[]{1,2}));
Console.WriteLine(Check(new[]{1,2,3}));"#
        ),
        &["empty", "single", "pair", "many"]
    );
}

#[test]
fn relational_and_pattern_combines_bounds_check() {
    assert_eq!(
        run_csharp(
            r#"string Grade(int n)=>n switch{
    >=90=>"A",
    >=70 and <90=>"B",
    >=50 and <70=>"C",
    _=>"F"};
Console.WriteLine(Grade(95));
Console.WriteLine(Grade(75));
Console.WriteLine(Grade(55));
Console.WriteLine(Grade(30));"#
        ),
        &["A", "B", "C", "F"]
    );
}

#[test]
fn type_pattern_in_switch_expression_dispatches_to_subclass() {
    assert_eq!(
        run_csharp(
            r#"abstract class Expr{}
class Num:Expr{public int V;}
class Add:Expr{public Expr L,R;}
int Eval(Expr e)=>e switch{
    Num n=>n.V,
    Add a=>Eval(a.L)+Eval(a.R),
    _=>throw new System.Exception()};
var tree=new Add{L=new Num{V=3},R=new Add{L=new Num{V=4},R=new Num{V=5}}};
Console.WriteLine(Eval(tree));"#
        ),
        &["12"]
    );
}

#[test]
fn or_pattern_matches_one_of_several_values() {
    assert_eq!(
        run_csharp(
            r#"string Weekend(string day)=>day switch{
    "Saturday" or "Sunday"=>"weekend",
    _=>"weekday"};
Console.WriteLine(Weekend("Saturday"));
Console.WriteLine(Weekend("Monday"));"#
        ),
        &["weekend", "weekday"]
    );
}
