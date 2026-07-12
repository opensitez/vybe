//! Constructor patterns: chaining, copy constructors, static factory, primary.
use super::helpers::run_csharp;

#[test]
fn constructor_chaining_via_this_keyword() {
    assert_eq!(
        run_csharp(
            r#"class Box{public int W,H,D;
    public Box(int w,int h,int d){W=w;H=h;D=d;}
    public Box(int side):this(side,side,side){}
}
var cube=new Box(3);
Console.WriteLine(cube.W); Console.WriteLine(cube.H); Console.WriteLine(cube.D);"#
        ),
        &["3", "3", "3"]
    );
}

#[test]
fn static_factory_method_creates_instance() {
    assert_eq!(
        run_csharp(
            r#"class Color{
    public int R,G,B;
    public static Color FromGray(int v)=>new Color{R=v,G=v,B=v};
}
var gray=Color.FromGray(128);
Console.WriteLine(gray.R==gray.G&&gray.G==gray.B);"#
        ),
        &["True"]
    );
}

#[test]
fn base_constructor_called_before_derived_body() {
    assert_eq!(
        run_csharp(
            r#"class A{public int Order;public A(){Order=1;}}
class B:A{public int Extra;public B():base(){Extra=2;}}
var b=new B();
Console.WriteLine(b.Order); Console.WriteLine(b.Extra);"#
        ),
        &["1", "2"]
    );
}

#[test]
fn multiple_constructors_via_overloading() {
    assert_eq!(
        run_csharp(
            r#"class Range{public int Lo,Hi;
    public Range():this(0,100){}
    public Range(int lo,int hi){Lo=lo;Hi=hi;}
}
var r1=new Range(); var r2=new Range(5,10);
Console.WriteLine(r1.Lo); Console.WriteLine(r2.Hi);"#
        ),
        &["0", "10"]
    );
}

#[test]
fn parameterless_constructor_required_for_generic_new_constraint() {
    assert_eq!(
        run_csharp(
            r#"class Widget{public int Value=7;}
T Make<T>() where T:new()=>new T();
Console.WriteLine(Make<Widget>().Value);"#
        ),
        &["7"]
    );
}

#[test]
fn primary_constructor_on_record_sets_all_fields() {
    assert_eq!(
        run_csharp(
            r#"record Person(string Name,int Age);
var p=new Person("Grace",40);
Console.WriteLine(p.Name); Console.WriteLine(p.Age);"#
        ),
        &["Grace", "40"]
    );
}
