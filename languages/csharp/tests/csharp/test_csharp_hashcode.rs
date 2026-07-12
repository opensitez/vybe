//! `System.HashCode` struct: Combine, Add, and custom GetHashCode patterns.
use super::helpers::run_csharp;

#[test]
fn hashcode_combine_two_values_produces_stable_result() {
    assert_eq!(
        run_csharp(
            r#"int h1=System.HashCode.Combine(1,2);
int h2=System.HashCode.Combine(1,2);
Console.WriteLine(h1==h2);"#
        ),
        &["True"]
    );
}

#[test]
fn hashcode_combine_different_values_differ() {
    assert_eq!(
        run_csharp(
            r#"int h1=System.HashCode.Combine(1,2);
int h2=System.HashCode.Combine(2,1);
Console.WriteLine(h1!=h2);"#
        ),
        &["True"]
    );
}

#[test]
fn hashcode_add_produces_same_as_combine_for_two_values() {
    assert_eq!(
        run_csharp(
            r#"var hc=new System.HashCode();
hc.Add(1); hc.Add(2);
int h1=hc.ToHashCode();
int h2=System.HashCode.Combine(1,2);
Console.WriteLine(h1==h2);"#
        ),
        &["True"]
    );
}

#[test]
fn class_overriding_get_hash_code_uses_hashcode_combine() {
    assert_eq!(
        run_csharp(
            r#"class Point{
    public int X,Y;
    public override int GetHashCode()=>System.HashCode.Combine(X,Y);
}
var p1=new Point{X=1,Y=2};
var p2=new Point{X=1,Y=2};
Console.WriteLine(p1.GetHashCode()==p2.GetHashCode());"#
        ),
        &["True"]
    );
}
