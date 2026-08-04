// vybe-test: csharp/csharp_design_patterns/singleton_returns_same_instance_on_repeated_calls
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

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

class Singleton{
    static Singleton _inst;
    public int Val;
    public static Singleton Instance=>_inst??=new Singleton();
}
Singleton.Instance.Val=42;
__P((Singleton.Instance.Val).ToString());
__Check("42");
