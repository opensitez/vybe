// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_int_adds_new_behaviour
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

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

static class IntExt { public static bool IsEven(this int n) => n%2==0; }
__P((4.IsEven()).ToString()); __P((3.IsEven()).ToString());
__Check("True\nFalse");
