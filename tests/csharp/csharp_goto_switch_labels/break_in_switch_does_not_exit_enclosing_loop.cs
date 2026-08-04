// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_does_not_exit_enclosing_loop
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

int i = 0;
while (i < 2) {
    switch (i) {
        case 0: i++; break;
        case 1: i++; break;
    }
}
__P((i).ToString());
__Check("2");
