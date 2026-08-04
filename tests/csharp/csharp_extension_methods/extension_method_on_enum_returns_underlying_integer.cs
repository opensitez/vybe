// vybe-test: csharp/csharp_extension_methods/extension_method_on_enum_returns_underlying_integer
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

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

using Demo; enum Mode { Off = 0, On = 2 } namespace Demo { public static class ModeExt { public static int Code(this Mode mode) { return (int)mode; } } } __P((Mode.On.Code()).ToString());
__Check("2");
