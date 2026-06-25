//! Overrides may return a more derived type than the virtual member (covariant returns).
use super::helpers::run_csharp;

#[test]
fn override_with_derived_return_type_is_callable_through_base_signature() {
    assert_eq!(
        run_csharp(
            r#"
class Animal { public string Name = "generic"; }
class Dog : Animal { }
class Shelter {
    public virtual Animal Adopt() { return new Animal(); }
}
class DogShelter : Shelter {
    public override Dog Adopt() { return new Dog { Name = "rex" }; }
}
Shelter place = new DogShelter();
Console.WriteLine(place.Adopt().Name);
"#
        ),
        &["rex"]
    );
}

#[test]
fn property_override_can_narrow_getter_return_type_covariantly() {
    assert_eq!(
        run_csharp(
            r#"
class Shape { public virtual Shape CloneShape() { return new Shape(); } }
class Circle : Shape { public int Radius; public override Circle CloneShape() { return new Circle { Radius = Radius }; } }
Shape original = new Circle { Radius = 5 };
var copy = original.CloneShape();
Console.WriteLine(copy is Circle);
"#
        ),
        &["True"]
    );
}
