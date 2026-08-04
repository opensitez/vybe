// vybe-test: csharp/interfaces_generics/extension_method_on_int
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

static class IntExtensions {
    public static bool IsEven(this int n) { return n % 2 == 0; }
    public static int Square(this int n) { return n * n; }
}
__P((4.IsEven()).ToString());
__P((3.IsEven()).ToString());
__P((5.Square()).ToString());
__Check("True\nFalse\n25");
