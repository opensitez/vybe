// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

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

// tuple_projection_checks
string feature = "tuple_projection_checks:36"; __P((feature.Length >= 1).ToString());
__Check("True");
