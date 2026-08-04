// vybe-test: csharp/csharp_static_classes/static_constructor_runs_once_before_first_member_access
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

class Singleton {
    public static int InitCount = 0;
    static Singleton() { InitCount++; }
    public static int Value = 42;
}
__P((Singleton.Value).ToString());
__P((Singleton.InitCount).ToString());
__Check("42\n1");
