// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
string feature = "constructor_overload_surface"; __Check((feature[0] == feature[0]).ToString(), "True");
