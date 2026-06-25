//! User-defined operator overloads: arithmetic, comparison, conversion.
use super::helpers::run_csharp;

#[test]
fn plus_operator_adds_two_vectors() {
    assert_eq!(
        run_csharp(r#"struct Vec{public int X,Y;
public static Vec operator+(Vec a,Vec b)=>new Vec{X=a.X+b.X,Y=a.Y+b.Y};}
var v=new Vec{X=1,Y=2}+new Vec{X=3,Y=4};
Console.WriteLine(v.X); Console.WriteLine(v.Y);"#),
        &["4", "6"]
    );
}

#[test]
fn minus_operator_subtracts_two_vectors() {
    assert_eq!(
        run_csharp(r#"struct Vec{public int X,Y;
public static Vec operator-(Vec a,Vec b)=>new Vec{X=a.X-b.X,Y=a.Y-b.Y};}
var v=new Vec{X=5,Y=3}-new Vec{X=2,Y=1};
Console.WriteLine(v.X);"#),
        &["3"]
    );
}

#[test]
fn unary_negation_operator_flips_sign_of_fields() {
    assert_eq!(
        run_csharp(r#"struct Vec{public int X;
public static Vec operator-(Vec v)=>new Vec{X=-v.X};}
var v=-new Vec{X=7};
Console.WriteLine(v.X);"#),
        &["-7"]
    );
}

#[test]
fn equality_operator_compares_value_type_fields() {
    assert_eq!(
        run_csharp(r#"struct Color{public int R,G,B;
public static bool operator==(Color a,Color b)=>a.R==b.R&&a.G==b.G&&a.B==b.B;
public static bool operator!=(Color a,Color b)=>!(a==b);
public override int GetHashCode()=>0; public override bool Equals(object o)=>o is Color c&&c==this;}
var a=new Color{R=1,G=2,B=3}; var b=new Color{R=1,G=2,B=3};
Console.WriteLine(a==b); Console.WriteLine(a!=b);"#),
        &["True", "False"]
    );
}

#[test]
fn comparison_operators_less_and_greater() {
    assert_eq!(
        run_csharp(r#"class Weight:System.IComparable<Weight>{
public int Kg;
public static bool operator<(Weight a,Weight b)=>a.Kg<b.Kg;
public static bool operator>(Weight a,Weight b)=>a.Kg>b.Kg;
public int CompareTo(Weight o)=>Kg.CompareTo(o.Kg);}
var a=new Weight{Kg=5}; var b=new Weight{Kg=10};
Console.WriteLine(a<b); Console.WriteLine(a>b);"#),
        &["True", "False"]
    );
}

#[test]
fn implicit_conversion_from_int_to_custom_type() {
    assert_eq!(
        run_csharp(r#"struct Meters{public double Value;
public static implicit operator Meters(double d)=>new Meters{Value=d};}
Meters m=3.5;
Console.WriteLine(m.Value);"#),
        &["3.5"]
    );
}

#[test]
fn explicit_conversion_to_primitive_requires_cast() {
    assert_eq!(
        run_csharp(r#"struct Percent{public double Value;
public static explicit operator double(Percent p)=>p.Value/100.0;}
var p=new Percent{Value=50};
Console.WriteLine((double)p);"#),
        &["0.5"]
    );
}

#[test]
fn multiply_operator_scales_vector_by_scalar() {
    assert_eq!(
        run_csharp(r#"struct Vec{public int X;
public static Vec operator*(Vec v,int s)=>new Vec{X=v.X*s};}
Console.WriteLine((new Vec{X=3}*4).X);"#),
        &["12"]
    );
}
