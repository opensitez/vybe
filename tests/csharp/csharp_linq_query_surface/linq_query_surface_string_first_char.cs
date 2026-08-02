// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
string feature = "linq_query_surface"; __Check((feature[0] == feature[0]).ToString(), "True");
