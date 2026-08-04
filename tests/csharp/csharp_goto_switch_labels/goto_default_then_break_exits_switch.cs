// vybe-test: csharp/csharp_goto_switch_labels/goto_default_then_break_exits_switch
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

int n = 0;
string r = "";
switch (n) {
    case 0:
        goto default;
    default:
        r = "done";
        break;
}
__P((r).ToString());
__Check("done");
