// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_parameter_in_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

static int Read<T>(ref readonly T value) where T: struct { return value.ToString().Length; } int n=123; __P((Read(ref n)>0).ToString());
__Check("True");
