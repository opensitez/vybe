// vybe-test: csharp/csharp_method_overload_resolution/out_parameter_is_assigned_before_caller_observes_result
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

bool TryHalve(int input, out int half) {
    if (input % 2 != 0) {
        half = 0;
        return false;
    }
    half = input / 2;
    return true;
}
if (TryHalve(8, out var result)) {
    __P((result).ToString());
} else {
    __P(("fail").ToString());
}
__Check("4");
