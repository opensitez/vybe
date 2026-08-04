// vybe-test: csharp/csharp_generics/generic_method
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

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

class Utils {
    public static T Identity<T>(T value) { return value; }
}
__P((Utils.Identity<int>(42)).ToString());
__P((Utils.Identity<string>("hello")).ToString());
__Check("42\nhello");
