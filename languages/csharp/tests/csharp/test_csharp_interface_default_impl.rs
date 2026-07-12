//! Default interface method implementations, static interface members (C# 8/11).
use super::helpers::run_csharp;

#[test]
fn default_method_used_when_class_does_not_override() {
    assert_eq!(
        run_csharp(
            r#"interface ILogger{
    void Log(string msg)=>Console.WriteLine("[LOG] "+msg);
}
class App:ILogger{}
ILogger app=new App();
app.Log("hello");"#
        ),
        &["[LOG] hello"]
    );
}

#[test]
fn class_can_override_default_interface_method() {
    assert_eq!(
        run_csharp(
            r#"interface ILogger{void Log(string msg)=>Console.WriteLine("[LOG] "+msg);}
class SilentApp:ILogger{public void Log(string msg){}}
ILogger app=new SilentApp();
app.Log("hello");
Console.WriteLine("done");"#
        ),
        &["done"]
    );
}

#[test]
fn two_classes_same_interface_one_overrides_one_uses_default() {
    assert_eq!(
        run_csharp(
            r#"interface IFormat{string Format(int n)=>$"[{n}]";}
class A:IFormat{}
class B:IFormat{public string Format(int n)=>n.ToString();}
IFormat a=new A(); IFormat b=new B();
Console.WriteLine(a.Format(5));
Console.WriteLine(b.Format(5));"#
        ),
        &["[5]", "5"]
    );
}
