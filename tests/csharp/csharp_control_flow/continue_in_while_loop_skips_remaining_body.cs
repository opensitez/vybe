// vybe-test: csharp/csharp_control_flow/continue_in_while_loop_skips_remaining_body
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

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
while (n < 5) {
    n++;
    if (n % 2 == 0) continue;
    sum += n;
}
__P((sum).ToString());
__Check("9");
