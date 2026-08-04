// vybe-test: csharp/control_flow/break_in_loop
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

var result = 0;
        for (var i = 0; i < 100; i++) {
            if (i == 5) break;
            result = result + 1;
        }
        __P((result).ToString());
__Check("5");
