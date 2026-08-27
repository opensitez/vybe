// vybe-test: csharp/csharp_class_features/static_method_returns_instance_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

using static __Harness;

__P((Logger.For("app").Format("hello")).ToString());
__Check("[app] hello");

class Logger{
    string prefix;
    Logger(string p){prefix=p;}
    public static Logger For(string name)=>new Logger(name);
    public string Format(string m)=>$"[{prefix}] {m}";
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
