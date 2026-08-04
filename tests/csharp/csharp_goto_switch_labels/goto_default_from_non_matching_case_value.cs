// vybe-test: csharp/csharp_goto_switch_labels/goto_default_from_non_matching_case_value
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

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

int n = 99;
string label = "";
switch (n) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    default:
        label = "other";
        break;
}
__P((label).ToString());
__Check("other");
