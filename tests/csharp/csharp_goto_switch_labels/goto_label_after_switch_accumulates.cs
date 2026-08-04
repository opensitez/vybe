// vybe-test: csharp/csharp_goto_switch_labels/goto_label_after_switch_accumulates
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

int code = 2;
int acc = 0;
switch (code) {
    case 1: acc += 1; break;
    case 2: acc += 2; goto default;
    default: acc += 100; break;
}
__P((acc).ToString());
__Check("102");
