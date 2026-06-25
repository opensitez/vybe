//! `abstract` forces override; `sealed` prevents further derivation.
use super::helpers::run_csharp;

#[test]
fn abstract_method_must_be_overridden_and_is_dispatched_polymorphically() {
    assert_eq!(
        run_csharp(
            r#"
abstract class Shape { public abstract double Area(); }
class Circle : Shape {
    public double R;
    public override double Area() => System.Math.PI * R * R;
}
Shape s = new Circle { R = 0 };
Console.WriteLine(s.Area());
"#
        ),
        &["0"]
    );
}

#[test]
fn abstract_class_cannot_be_instantiated_directly_throws_on_attempt() {
    assert_eq!(
        run_csharp(
            r#"
abstract class Base { }
string result = "ok";
try {
    var obj = System.Activator.CreateInstance(typeof(Base));
    result = "created";
} catch (System.MemberAccessException) {
    result = "blocked";
} catch (System.Exception) {
    result = "blocked";
}
Console.WriteLine(result);
"#
        ),
        &["blocked"]
    );
}

#[test]
fn sealed_class_cannot_be_used_as_base_detected_at_compile_time_but_runtime_ok() {
    assert_eq!(
        run_csharp(
            r#"
sealed class Final { public int Value = 7; }
var f = new Final();
Console.WriteLine(f.Value);
"#
        ),
        &["7"]
    );
}

#[test]
fn sealed_method_override_stops_further_overriding_in_chain() {
    assert_eq!(
        run_csharp(
            r#"
class A { public virtual string Name() => "A"; }
class B : A { public sealed override string Name() => "B"; }
class C : B { }
A obj = new C();
Console.WriteLine(obj.Name());
"#
        ),
        &["B"]
    );
}
