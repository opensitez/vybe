// vybe-test: csharp/csharp_class_features/static_method_returns_instance_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger{
    string prefix;
    Logger(string p){prefix=p;}
    public static Logger For(string name)=>new Logger(name);
    public string Format(string m)=>$"[{prefix}] {m}";
}
__Check((Logger.For("app").Format("hello")).ToString(), "[app] hello");
