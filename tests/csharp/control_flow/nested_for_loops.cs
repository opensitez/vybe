// vybe-test: csharp/control_flow/nested_for_loops
// origin: languages/csharp/tests/csharp/test_control_flow.rs

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

var sum = 0;
        for (var i = 0; i < 3; i++) {
            for (var j = 0; j < 3; j++) {
                sum = sum + 1;
            }
        }
        __P((sum).ToString());
__Check("9");
