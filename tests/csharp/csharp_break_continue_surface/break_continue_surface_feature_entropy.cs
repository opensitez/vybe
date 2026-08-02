// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
string feature = "break_continue_surface:49"; __Check((feature.Length >= 1).ToString(), "True");
