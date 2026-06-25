//! Explicit interface implementation hides members behind the interface reference.
use super::helpers::run_csharp;

#[test]
fn explicit_impl_is_not_accessible_through_class_reference() {
    assert_eq!(
        run_csharp(
            r#"
interface IArea { double Area(); }
class Square : IArea {
    public double Side;
    double IArea.Area() => Side * Side;
}
IArea shape = new Square { Side = 3 };
Console.WriteLine(shape.Area());
"#
        ),
        &["9"]
    );
}

#[test]
fn class_with_both_explicit_and_public_overloads_picks_by_static_type() {
    assert_eq!(
        run_csharp(
            r#"
interface IDescribe { string Describe(); }
class Widget : IDescribe {
    public string Describe() => "widget";
    string IDescribe.Describe() => "interface:widget";
}
var w = new Widget();
IDescribe i = w;
Console.WriteLine(w.Describe());
Console.WriteLine(i.Describe());
"#
        ),
        &["widget", "interface:widget"]
    );
}

#[test]
fn two_interfaces_with_same_method_name_implemented_explicitly_route_separately() {
    assert_eq!(
        run_csharp(
            r#"
interface ILeft  { string Side(); }
interface IRight { string Side(); }
class Both : ILeft, IRight {
    string ILeft.Side()  => "left";
    string IRight.Side() => "right";
}
ILeft  l = new Both();
IRight r = new Both();
Console.WriteLine(l.Side());
Console.WriteLine(r.Side());
"#
        ),
        &["left", "right"]
    );
}
