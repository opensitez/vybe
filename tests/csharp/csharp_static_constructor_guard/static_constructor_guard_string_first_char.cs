// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
string feature = "static_constructor_guard"; __Check((feature[0] == feature[0]).ToString(), "True");
