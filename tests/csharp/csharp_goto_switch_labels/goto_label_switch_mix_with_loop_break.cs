// vybe-test: csharp/csharp_goto_switch_labels/goto_label_switch_mix_with_loop_break
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

string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            log += "0";
            break;
        case 1:
            log += "1";
            break;
    }
    if (i == 1) break;
}
__P((log).ToString());
__Check("01");
