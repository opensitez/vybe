// vybe-test: csharp/csharp_convert_uri_path/convert_to_int32_truncates_double_toward_zero
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// convert_uri_path
__Check((System.Convert.ToInt32(3.9)).ToString(), "3");
