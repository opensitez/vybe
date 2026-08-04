// vybe-test: csharp/csharp_local_functions_partial_methods/static_local_function_does_not_capture_outer_state
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

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

static int Triple(int value) { return value * 3; } __P((Triple(4)).ToString());
__Check("12");
