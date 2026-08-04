// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_class_allows_reference_type_members
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

string Describe<T>(T value) where T : class {
    return value == null ? "null" : value.ToString();
}
__P((Describe("data")).ToString());
__Check("data");
