// vybe-test: csharp/csharp_static_classes/static_method_can_call_other_static_methods_in_same_class
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

static class Calc {
    static int Add(int a, int b) => a+b;
    public static int Sum3(int a, int b, int c) => Add(Add(a,b),c);
}
__P((Calc.Sum3(1,2,3)).ToString());
__Check("6");
