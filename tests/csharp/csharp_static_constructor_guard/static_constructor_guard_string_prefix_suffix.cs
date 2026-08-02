// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
string feature = "static_constructor_guard"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
