// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
string feature = "serialization_json_surface:91"; __Check((feature.Length >= 1).ToString(), "True");
