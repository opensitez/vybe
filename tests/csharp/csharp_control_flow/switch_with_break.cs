// vybe-test: csharp/csharp_control_flow/switch_with_break
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

int day = 3;
switch (day) {
    case 1: __P(("Mon").ToString()); break;
    case 2: __P(("Tue").ToString()); break;
    case 3: __P(("Wed").ToString()); break;
    default: __P(("Other").ToString()); break;
}
__Check("Wed");
