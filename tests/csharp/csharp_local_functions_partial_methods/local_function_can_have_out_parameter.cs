// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_have_out_parameter
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

void Split(int value, out int left, out int right) { left = value / 2; right = value - left; } Split(9, out var left, out var right); __P((left).ToString()); __P((right).ToString());
__Check("4\n5");
