// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_struct_accepts_unboxed_value_type
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

int Scale<T>(T value) where T : struct {
    return 2 * (int)(object)value;
}
__P((Scale(6)).ToString());
__Check("12");
