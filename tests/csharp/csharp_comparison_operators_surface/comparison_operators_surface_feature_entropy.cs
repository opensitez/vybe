// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

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

// comparison_operators_surface
string feature = "comparison_operators_surface:13"; __P((feature.Length >= 1).ToString());
__Check("True");
