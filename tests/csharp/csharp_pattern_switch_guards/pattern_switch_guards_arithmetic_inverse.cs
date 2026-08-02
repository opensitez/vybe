// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
int seed = 42; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
