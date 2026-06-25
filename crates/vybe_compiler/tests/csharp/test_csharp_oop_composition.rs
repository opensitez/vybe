//! OOP composition over inheritance: delegation, decorator, mixin-like patterns.
use super::helpers::run_csharp;

#[test]
fn composition_delegates_to_contained_object() {
    assert_eq!(
        run_csharp(r#"class Logger{public void Log(string m)=>Console.WriteLine("[LOG]"+m);}
class Service{
    readonly Logger _log=new Logger();
    public void Do(string m){_log.Log(m);}
}
new Service().Do("hello");"#),
        &["[LOG]hello"]
    );
}

#[test]
fn decorator_wraps_and_extends_base_behaviour() {
    assert_eq!(
        run_csharp(r#"interface IText{string Get();}
class Plain:IText{public string Get()=>"hello";}
class Shout:IText{
    IText _inner;
    public Shout(IText inner){_inner=inner;}
    public string Get()=>_inner.Get().ToUpper()+"!";
}
IText t=new Shout(new Plain());
Console.WriteLine(t.Get());"#),
        &["HELLO!"]
    );
}

#[test]
fn chained_decorators_apply_in_order() {
    assert_eq!(
        run_csharp(r#"interface IText{string Get();}
class Plain:IText{public string Get()=>"hello";}
class Shout:IText{IText i;public Shout(IText x){i=x;}public string Get()=>i.Get().ToUpper();}
class Wrap:IText{IText i;public Wrap(IText x){i=x;}public string Get()=>$"[{i.Get()}]";}
IText t=new Wrap(new Shout(new Plain()));
Console.WriteLine(t.Get());"#),
        &["[HELLO]"]
    );
}

#[test]
fn strategy_injected_via_constructor_delegation() {
    assert_eq!(
        run_csharp(r#"interface IFormatter{string Format(int n);}
class Hex:IFormatter{public string Format(int n)=>n.ToString("X");}
class Dec:IFormatter{public string Format(int n)=>n.ToString();}
class Printer{
    IFormatter _f;
    public Printer(IFormatter f){_f=f;}
    public string Print(int n)=>_f.Format(n);
}
Console.WriteLine(new Printer(new Hex()).Print(255));
Console.WriteLine(new Printer(new Dec()).Print(255));"#),
        &["FF", "255"]
    );
}
