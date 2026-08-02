// vybe-test: csharp/csharp_convert_uri_path/convert_to_boolean_maps_nonzero_integers_to_true
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// convert_uri_path
__Check((System.Convert.ToBoolean(1)).ToString(), "True");
