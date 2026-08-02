// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
string feature = "for_loop_bounds"; __Check((feature.Length > 0).ToString(), "True");
