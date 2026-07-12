//! `typeof` on a type token differs from `GetType()` on a polymorphic instance.
use super::helpers::run_csharp;

#[test]
fn typeof_reports_declared_type_while_gettype_reports_runtime_type_of_instance() {
    assert_eq!(
        run_csharp(
            r#"
class Animal { }
class Dog : Animal { }
Animal pet = new Dog();
Console.WriteLine(typeof(Animal).Name);
Console.WriteLine(pet.GetType().Name);
"#
        ),
        &["Animal", "Dog"]
    );
}
