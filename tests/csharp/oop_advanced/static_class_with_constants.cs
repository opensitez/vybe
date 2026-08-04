// vybe-test: csharp/oop_advanced/static_class_with_constants
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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

static class Constants {
    public const double Pi = 3.14159;
    public const int MaxSize = 100;
}
__P((Constants.Pi).ToString());
__P((Constants.MaxSize).ToString());
__Check("3.14159\n100");
