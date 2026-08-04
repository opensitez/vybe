// vybe-test: csharp/csharp_control_flow/nested_loops
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

int count = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 4; j++) {
        count++;
    }
}
__P((count).ToString());
__Check("12");
