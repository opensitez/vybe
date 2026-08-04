// vybe-test: csharp/type_features/csharp_uses_host_namespace
// origin: languages/csharp/tests/csharp/test_type_features.rs

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

__P((Math.Floor(9.7)).ToString());
        __P((Math.Abs(-42)).ToString());
        __P((Math.Sqrt(144)).ToString());
__Check("9\n42\n12");
