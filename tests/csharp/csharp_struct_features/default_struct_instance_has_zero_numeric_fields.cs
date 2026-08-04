// vybe-test: csharp/csharp_struct_features/default_struct_instance_has_zero_numeric_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

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

struct Size { public int W, H; }
Size s = default;
__P((s.W).ToString()); __P((s.H).ToString());
__Check("0\n0");
