// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
string feature = "static_constructor_guard:69"; __Check((feature.Length >= 1).ToString(), "True");
