// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_calls_method_on_expression_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

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

var text = "hi"; __P(($"{text.ToUpper()}").ToString());
__Check("HI");
