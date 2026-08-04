// vybe-test: csharp/csharp_goto_switch_labels/switch_fallthrough_via_goto_case_only
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

int v = 1;
int total = 0;
switch (v) {
    case 1: total += 10; goto case 2;
    case 2: total += 1; break;
}
__P((total).ToString());
__Check("11");
