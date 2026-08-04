// vybe-test: csharp/csharp_nested_control_flow/do_while_executes_body_before_first_condition_test
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

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

int count = 0;
do {
    count++;
} while (count < 1);
__P((count).ToString());
__Check("1");
