// vybe-test: csharp/csharp_enum_flags_operations/enum_switch_dispatches_on_underlying_constant_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

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

enum Mode { Alpha = 1, Beta = 2 }
string Label(Mode mode) {
    switch (mode) {
        case Mode.Alpha: return "a";
        case Mode.Beta: return "b";
        default: return "?";
    }
}
__P((Label(Mode.Beta)).ToString());
__Check("b");
