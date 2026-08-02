// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
double seed = 69; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
