// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
string feature = "break_continue_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
