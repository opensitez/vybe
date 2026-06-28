//! Advanced record features: inheritance, custom methods, record struct, Deconstruct.
use super::helpers::run_csharp;

#[test]
fn derived_record_inherits_base_record_properties() {
    assert_eq!(
        run_csharp(
            r#"record Animal(string Name);
record Dog(string Name,string Breed):Animal(Name);
var d=new Dog("Rex","Lab");
Console.WriteLine(d.Name); Console.WriteLine(d.Breed);"#
        ),
        &["Rex", "Lab"]
    );
}

#[test]
fn record_with_expression_creates_shallow_copy_with_changes() {
    assert_eq!(
        run_csharp(
            r#"record Config(int Port,string Host);
var c1=new Config(80,"localhost");
var c2=c1 with{Port=443};
Console.WriteLine(c1.Port); Console.WriteLine(c2.Port);
Console.WriteLine(c2.Host);"#
        ),
        &["80", "443", "localhost"]
    );
}

#[test]
fn record_deconstruct_works_in_foreach() {
    assert_eq!(
        run_csharp(
            r#"record Point(int X,int Y);
var pts=new[]{new Point(1,2),new Point(3,4)};
int sumX=0;
foreach(var(x,_) in pts) sumX+=x;
Console.WriteLine(sumX);"#
        ),
        &["4"]
    );
}

#[test]
fn record_struct_has_value_semantics() {
    assert_eq!(
        run_csharp(
            r#"record struct Vec(int X,int Y);
var a=new Vec(1,2); var b=a; // copy
b=b with{X=99};
Console.WriteLine(a.X);"#
        ),
        &["1"]
    );
}

#[test]
fn record_equals_compares_all_properties_by_value() {
    assert_eq!(
        run_csharp(
            r#"record Pair(int A,int B);
var p1=new Pair(1,2); var p2=new Pair(1,2); var p3=new Pair(1,3);
Console.WriteLine(p1==p2);
Console.WriteLine(p1==p3);"#
        ),
        &["True", "False"]
    );
}

#[test]
fn record_custom_method_works_alongside_generated_members() {
    assert_eq!(
        run_csharp(
            r#"record Circle(double Radius){
    public double Area=>System.Math.PI*Radius*Radius;
}
var c=new Circle(1.0);
Console.WriteLine(c.Area>3.1&&c.Area<3.2);"#
        ),
        &["True"]
    );
}
