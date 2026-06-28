//! Nested class declarations, access rules, and inner-outer interaction.
use super::helpers::run_csharp;

#[test]
fn nested_class_can_access_outer_private_members() {
    assert_eq!(
        run_csharp(
            r#"class Outer{
    static int secret=42;
    public class Inner{public int Get()=>secret;}
}
Console.WriteLine(new Outer.Inner().Get());"#
        ),
        &["42"]
    );
}

#[test]
fn nested_class_fully_qualified_via_outer_name() {
    assert_eq!(
        run_csharp(
            r#"class Container{public class Item{public int Value=7;}}
var item=new Container.Item();
Console.WriteLine(item.Value);"#
        ),
        &["7"]
    );
}

#[test]
fn nested_static_class_provides_utility_methods() {
    assert_eq!(
        run_csharp(
            r#"class Parser{
    public static class Helpers{public static int ToInt(string s)=>int.Parse(s);}
}
Console.WriteLine(Parser.Helpers.ToInt("99"));"#
        ),
        &["99"]
    );
}

#[test]
fn deeply_nested_class_visible_through_chain() {
    assert_eq!(
        run_csharp(
            r#"class A{public class B{public class C{public int V=3;}}}
Console.WriteLine(new A.B.C().V);"#
        ),
        &["3"]
    );
}

#[test]
fn nested_class_can_implement_interface() {
    assert_eq!(
        run_csharp(
            r#"interface IValue{int Get();}
class Host{public class Impl:IValue{public int Get()=>5;}}
IValue v=new Host.Impl();
Console.WriteLine(v.Get());"#
        ),
        &["5"]
    );
}

#[test]
fn nested_class_inherits_from_another_nested_class() {
    assert_eq!(
        run_csharp(
            r#"class Shapes{
    public class Base{public virtual string Name()=>"shape";}
    public class Circle:Base{public override string Name()=>"circle";}
}
Shapes.Base b=new Shapes.Circle();
Console.WriteLine(b.Name());"#
        ),
        &["circle"]
    );
}
