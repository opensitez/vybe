// vybe-test: csharp/csharp_method_group_delegates/method_group_converts_to_func_without_explicit_lambda_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_method_group_delegates.rs

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

static int Double(int n) => n * 2;
System.Func<int, int> fn = Double;
__P((fn(6)).ToString());
__Check("12");
