// vybe-test: csharp/csharp_checked_context_math/checked_context_math_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
string feature = "checked_context_math"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
