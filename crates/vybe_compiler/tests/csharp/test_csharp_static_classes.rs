//! Static classes, static constructors, and static method dispatch.
use super::helpers::run_csharp;

#[test]
fn static_class_method_callable_without_instance() {
    assert_eq!(
        run_csharp(
            r#"static class MathHelper { public static int Square(int n) => n*n; }
Console.WriteLine(MathHelper.Square(5));"#
        ),
        &["25"]
    );
}

#[test]
fn static_field_shared_across_all_callers() {
    assert_eq!(
        run_csharp(
            r#"static class Counter { public static int Count = 0; }
Counter.Count++;
Counter.Count++;
Console.WriteLine(Counter.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn static_constructor_runs_once_before_first_member_access() {
    assert_eq!(
        run_csharp(
            r#"class Singleton {
    public static int InitCount = 0;
    static Singleton() { InitCount++; }
    public static int Value = 42;
}
Console.WriteLine(Singleton.Value);
Console.WriteLine(Singleton.InitCount);"#
        ),
        &["42", "1"]
    );
}

#[test]
fn static_constructor_not_re_run_on_second_access() {
    assert_eq!(
        run_csharp(
            r#"class Registry {
    public static int Boot = 0;
    static Registry() { Boot++; }
    public static void Touch() { }
}
Registry.Touch();
Registry.Touch();
Console.WriteLine(Registry.Boot);"#
        ),
        &["1"]
    );
}

#[test]
fn static_readonly_field_set_in_static_constructor() {
    assert_eq!(
        run_csharp(
            r#"class Config {
    public static readonly string Version;
    static Config() { Version = "1.0"; }
}
Console.WriteLine(Config.Version);"#
        ),
        &["1.0"]
    );
}

#[test]
fn static_method_can_call_other_static_methods_in_same_class() {
    assert_eq!(
        run_csharp(
            r#"static class Calc {
    static int Add(int a, int b) => a+b;
    public static int Sum3(int a, int b, int c) => Add(Add(a,b),c);
}
Console.WriteLine(Calc.Sum3(1,2,3));"#
        ),
        &["6"]
    );
}
