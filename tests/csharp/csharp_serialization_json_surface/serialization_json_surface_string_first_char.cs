// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
string feature = "serialization_json_surface"; __Check((feature[0] == feature[0]).ToString(), "True");
