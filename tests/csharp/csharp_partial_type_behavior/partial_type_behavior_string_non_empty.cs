// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

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

// partial_type_behavior
string feature = "partial_type_behavior"; __P((feature.Length > 0).ToString());
__Check("True");
