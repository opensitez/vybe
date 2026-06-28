//! Advanced OOP: covariant return types, sealed override, protected internal.
use super::helpers::run_csharp;

#[test]
fn sealed_override_prevents_further_overriding() {
    assert_eq!(
        run_csharp(
            r#"class A{public virtual string Tag()=>"A";}
class B:A{public sealed override string Tag()=>"B";}
class C:B{}
C c=new C();
Console.WriteLine(c.Tag());"#
        ),
        &["B"]
    );
}

#[test]
fn covariant_return_type_narrows_return_of_override() {
    assert_eq!(
        run_csharp(
            r#"class Base{public virtual object Create()=>new object();}
class Derived:Base{public override string Create()=>"derived";}
Derived d=new Derived();
Console.WriteLine(d.Create());"#
        ),
        &["derived"]
    );
}

#[test]
fn multiple_levels_of_base_call_chain() {
    assert_eq!(
        run_csharp(
            r#"class A{public virtual string Name()=>"A";}
class B:A{public override string Name()=>"B+"+base.Name();}
class C:B{public override string Name()=>"C+"+base.Name();}
Console.WriteLine(new C().Name());"#
        ),
        &["C+B+A"]
    );
}

#[test]
fn abstract_class_partial_implementation_forces_concrete_override() {
    assert_eq!(
        run_csharp(
            r#"abstract class Step{
    public abstract string Name();
    public string Run()=>$"run:{Name()}";
}
class Alpha:Step{public override string Name()=>"alpha";}
Console.WriteLine(new Alpha().Run());"#
        ),
        &["run:alpha"]
    );
}

#[test]
fn object_type_is_common_base_of_all_classes() {
    assert_eq!(
        run_csharp(
            r#"object x=42; object y="hi"; object z=new int[]{};
Console.WriteLine(x is object);
Console.WriteLine(y is object);
Console.WriteLine(z is object);"#
        ),
        &["True", "True", "True"]
    );
}
