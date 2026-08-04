// vybe-test: csharp/csharp_generic_inference_calls/generic_method_overload_resolution_prefers_specific_argument_types
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

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

string Pick(int value) { return "int:" + value; }
string Pick(string value) { return "str:" + value; }
__P((Pick(3)).ToString());
__P((Pick("3")).ToString());
__Check("int:3\nstr:3");
