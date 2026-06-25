//! Expression-bodied members: methods, properties, constructors, indexers.
use super::helpers::run_csharp;

#[test]
fn expression_bodied_method_returns_computed_value() {
    assert_eq!(
        run_csharp(r#"class Calc{public int Double(int n)=>n*2;}
Console.WriteLine(new Calc().Double(5));"#),
        &["10"]
    );
}

#[test]
fn expression_bodied_property_getter() {
    assert_eq!(
        run_csharp(r#"class Circle{public double R;public double Area=>System.Math.PI*R*R;}
Console.WriteLine(System.Math.Round(new Circle{R=0}.Area));"#),
        &["0"]
    );
}

#[test]
fn expression_bodied_constructor() {
    assert_eq!(
        run_csharp(r#"class Point{public int X,Y; public Point(int x,int y)=>(X,Y)=(x,y);}
var p=new Point(3,4);
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#),
        &["3", "4"]
    );
}

#[test]
fn expression_bodied_indexer() {
    assert_eq!(
        run_csharp(r#"class Bag{int[]data={1,2,3};public int this[int i]=>data[i];}
Console.WriteLine(new Bag()[2]);"#),
        &["3"]
    );
}

#[test]
fn expression_bodied_static_method() {
    assert_eq!(
        run_csharp(r#"static class Utils{public static int Clamp(int v,int lo,int hi)=>v<lo?lo:v>hi?hi:v;}
Console.WriteLine(Utils.Clamp(15,0,10));"#),
        &["10"]
    );
}

#[test]
fn expression_bodied_void_method_using_statement_form() {
    assert_eq!(
        run_csharp(r#"class Logger{public void Log(string msg)=>Console.WriteLine(msg);}
new Logger().Log("hello");"#),
        &["hello"]
    );
}

#[test]
fn expression_bodied_operator() {
    assert_eq!(
        run_csharp(r#"struct Num{public int V;public static Num operator+(Num a,Num b)=>new Num{V=a.V+b.V};}
Console.WriteLine((new Num{V=3}+new Num{V=4}).V);"#),
        &["7"]
    );
}
