// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
int seed = 69; __Check((seed - seed == 0).ToString(), "True");
