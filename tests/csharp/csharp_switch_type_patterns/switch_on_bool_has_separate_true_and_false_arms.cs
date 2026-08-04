// vybe-test: csharp/csharp_switch_type_patterns/switch_on_bool_has_separate_true_and_false_arms
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

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

bool ok = false;
string label = ok switch { true => "yes", false => "no" };
__P((label).ToString());
__Check("no");
