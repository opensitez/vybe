// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_inside_loop_runs_once
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

string report = "";
for (int i = 0; i < 3; i++) {
    switch (i) {
        case 0: report += "a"; break;
        case 1: report += "b"; break;
        case 2: report += "c"; break;
    }
}
__P((report).ToString());
__Check("abc");
