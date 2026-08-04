// vybe-test: csharp/csharp_method_overload_resolution/params_array_overload_receives_remaining_arguments_as_array
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

int Sum(params int[] values) {
    int total = 0;
    foreach (var v in values) total += v;
    return total;
}
__P((Sum(1, 2, 3)).ToString());
__P((Sum()).ToString());
__Check("6\n0");
