// vybe-test: csharp/csharp_nested_control_flow/break_inside_inner_loop_does_not_stop_outer_loop
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

int total = 0;
for (int row = 0; row < 2; row++) {
    for (int col = 0; col < 4; col++) {
        if (col == 2) break;
        total += 1;
    }
}
__P((total).ToString());
__Check("4");
