// vybe-test: csharp/csharp_goto_switch_labels/continue_in_do_while_skips_body
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
int sum = 0;
do {
    n++;
    if (n == 2) continue;
    sum += n;
} while (n < 4);
__P((sum).ToString());
__Check("8");
