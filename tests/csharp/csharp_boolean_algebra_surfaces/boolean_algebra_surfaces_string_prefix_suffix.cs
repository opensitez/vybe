// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
