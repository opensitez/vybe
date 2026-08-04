// vybe-test: csharp/csharp_class_features/static_method_returns_instance_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Logger{
    string prefix;
    Logger(string p){prefix=p;}
    public static Logger For(string name)=>new Logger(name);
    public string Format(string m)=>$"[{prefix}] {m}";
}
__P((Logger.For("app").Format("hello")).ToString());
__Check("[app] hello");
