// vybe-test: csharp/csharp_switch_type_patterns/switch_on_int_falls_through_to_default_when_no_case_matches
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

int code = 99;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    default: label = "other"; break;
}
__P((label).ToString());
__Check("other");
