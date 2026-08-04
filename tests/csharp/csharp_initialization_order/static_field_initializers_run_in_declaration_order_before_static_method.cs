// vybe-test: csharp/csharp_initialization_order/static_field_initializers_run_in_declaration_order_before_static_method
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

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

class Logger {
    static string First = Mark("first");
    static string Second = Mark("second");
    static string Mark(string name) {
        __P((name).ToString());
        return name;
    }
    public static void Run() {
        __P(("run").ToString());
    }
}
Logger.Run();
__Check("first\nsecond\nrun");
