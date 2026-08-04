// vybe-test: csharp/csharp_static_classes/static_constructor_not_re_run_on_second_access
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

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

class Registry {
    public static int Boot = 0;
    static Registry() { Boot++; }
    public static void Touch() { }
}
Registry.Touch();
Registry.Touch();
__P((Registry.Boot).ToString());
__Check("1");
