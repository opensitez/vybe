// vybe-test: csharp/csharp_method_overload_resolution/optional_parameter_uses_default_when_argument_omitted
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

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

string FormatLine(string text, int level = 1) {
    return level + ":" + text;
}
__P((FormatLine("ok")).ToString());
__P((FormatLine("warn", 3)).ToString());
__Check("1:ok\n3:warn");
