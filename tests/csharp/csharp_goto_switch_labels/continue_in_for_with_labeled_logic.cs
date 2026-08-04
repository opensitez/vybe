// vybe-test: csharp/csharp_goto_switch_labels/continue_in_for_with_labeled_logic
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

string chars = "";
for (int i = 0; i < 4; i++) {
    if (i == 2) continue;
    chars += i.ToString();
}
__P((chars).ToString());
__Check("013");
