//! Advanced struct patterns: `ref struct`, `Span`-backed struct, operator overloads.
use super::helpers::run_csharp;

#[test]
fn struct_with_operator_overload_usable_in_expressions() {
    assert_eq!(
        run_csharp(
            r#"struct Fraction{
    public int Num,Den;
    public static Fraction operator+(Fraction a,Fraction b)=>
        new Fraction{Num=a.Num*b.Den+b.Num*a.Den,Den=a.Den*b.Den};
    public override string ToString()=>$"{Num}/{Den}";
}
var r=new Fraction{Num=1,Den=2}+new Fraction{Num=1,Den=3};
Console.WriteLine(r.Num); Console.WriteLine(r.Den);"#
        ),
        &["5", "6"]
    );
}

#[test]
fn struct_iequatable_implementation_compares_by_value() {
    assert_eq!(
        run_csharp(
            r#"struct Color:System.IEquatable<Color>{
    public int R,G,B;
    public bool Equals(Color o)=>R==o.R&&G==o.G&&B==o.B;
    public override bool Equals(object o)=>o is Color c&&Equals(c);
    public override int GetHashCode()=>System.HashCode.Combine(R,G,B);
}
var red1=new Color{R=255,G=0,B=0};
var red2=new Color{R=255,G=0,B=0};
Console.WriteLine(red1.Equals(red2));"#
        ),
        &["True"]
    );
}

#[test]
fn struct_default_keyword_produces_zero_fields() {
    assert_eq!(
        run_csharp(
            r#"struct Vec{public int X,Y,Z;}
var v=default(Vec);
Console.WriteLine(v.X==0&&v.Y==0&&v.Z==0);"#
        ),
        &["True"]
    );
}

#[test]
fn readonly_struct_method_cannot_modify_state() {
    assert_eq!(
        run_csharp(
            r#"readonly struct Counter{
    public readonly int Value;
    public Counter(int v){Value=v;}
    public Counter Increment()=>new Counter(Value+1);
}
var c=new Counter(5).Increment();
Console.WriteLine(c.Value);"#
        ),
        &["6"]
    );
}

#[test]
fn struct_passed_by_in_not_copied_but_read_only() {
    assert_eq!(
        run_csharp(
            r#"struct Vec{public int X,Y;}
int Sum(in Vec v)=>v.X+v.Y;
var v=new Vec{X=3,Y=4};
Console.WriteLine(Sum(in v));"#
        ),
        &["7"]
    );
}
