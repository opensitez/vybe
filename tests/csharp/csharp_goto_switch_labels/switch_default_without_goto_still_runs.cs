// vybe-test: csharp/csharp_goto_switch_labels/switch_default_without_goto_still_runs
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

int v = 5;
string tag = "";
switch (v) {
    case 1: tag = "one"; break;
    default: tag = "many"; break;
}
__P((tag).ToString());
__Check("many");
