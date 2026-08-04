// vybe-test: csharp/csharp_method_overload_resolution/more_specific_overload_wins_over_params_fallback_for_fixed_arity
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

string Describe(int value) { return "int:" + value; }
string Describe(params int[] values) { return "many:" + values.Length; }
__P((Describe(7)).ToString());
__P((Describe(1, 2)).ToString());
__Check("int:7\nmany:2");
