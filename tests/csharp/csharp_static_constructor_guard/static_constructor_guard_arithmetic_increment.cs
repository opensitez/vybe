// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
int seed = 69; __Check((seed + 1 > seed).ToString(), "True");
