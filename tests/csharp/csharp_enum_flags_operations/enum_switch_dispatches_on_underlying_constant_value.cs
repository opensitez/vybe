// vybe-test: csharp/csharp_enum_flags_operations/enum_switch_dispatches_on_underlying_constant_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((Label(Mode.Beta)).ToString(), "b");
