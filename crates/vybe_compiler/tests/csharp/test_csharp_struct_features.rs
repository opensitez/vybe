//! Value-type struct semantics, readonly structs, and default values.
use super::helpers::run_csharp;

#[test]
fn struct_assignment_copies_all_fields_independently() {
    assert_eq!(
        run_csharp(
            r#"struct Point { public int X, Y; }
var a = new Point { X=1, Y=2 };
var b = a;
b.X = 99;
Console.WriteLine(a.X);"#
        ),
        &["1"]
    );
}

#[test]
fn default_struct_instance_has_zero_numeric_fields() {
    assert_eq!(
        run_csharp(
            r#"struct Size { public int W, H; }
Size s = default;
Console.WriteLine(s.W); Console.WriteLine(s.H);"#
        ),
        &["0", "0"]
    );
}

#[test]
fn struct_with_custom_constructor_sets_fields() {
    assert_eq!(
        run_csharp(
            r#"struct Rect { public int W,H; public Rect(int w, int h) { W=w; H=h; } }
var r = new Rect(3, 4);
Console.WriteLine(r.W * r.H);"#
        ),
        &["12"]
    );
}

#[test]
fn readonly_struct_field_cannot_be_mutated_but_is_readable() {
    assert_eq!(
        run_csharp(
            r#"readonly struct Immutable { public readonly int Value; public Immutable(int v) { Value=v; } }
var obj = new Immutable(7);
Console.WriteLine(obj.Value);"#
        ),
        &["7"]
    );
}

#[test]
fn struct_method_operates_on_own_fields() {
    assert_eq!(
        run_csharp(
            r#"struct Vector { public double X,Y; public double Length() => System.Math.Sqrt(X*X+Y*Y); }
var v = new Vector { X=3, Y=4 };
Console.WriteLine(v.Length());"#
        ),
        &["5"]
    );
}

#[test]
fn passing_struct_to_method_copies_value() {
    assert_eq!(
        run_csharp(
            r#"struct Counter { public int N; }
void Increment(Counter c) { c.N++; }
var c = new Counter { N=5 };
Increment(c);
Console.WriteLine(c.N);"#
        ),
        &["5"]
    );
}

#[test]
fn struct_passed_by_ref_mutates_caller_copy() {
    assert_eq!(
        run_csharp(
            r#"struct Counter { public int N; }
void Increment(ref Counter c) { c.N++; }
var c = new Counter { N=5 };
Increment(ref c);
Console.WriteLine(c.N);"#
        ),
        &["6"]
    );
}

#[test]
fn struct_equality_via_overridden_equals() {
    assert_eq!(
        run_csharp(
            r#"struct Color {
    public int R,G,B;
    public override bool Equals(object o) => o is Color c && c.R==R && c.G==G && c.B==B;
    public override int GetHashCode() => System.HashCode.Combine(R,G,B);
}
var x = new Color{R=1,G=2,B=3};
var y = new Color{R=1,G=2,B=3};
Console.WriteLine(x.Equals(y));"#
        ),
        &["True"]
    );
}
