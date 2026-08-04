// vybe-test: csharp/csharp_nested_control_flow/continue_inside_inner_loop_skips_remaining_body_but_not_outer
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

int sum = 0;
for (int outer = 0; outer < 2; outer++) {
    for (int inner = 0; inner < 3; inner++) {
        if (inner == 1) continue;
        sum += inner;
    }
}
__P((sum).ToString());
__Check("4");
