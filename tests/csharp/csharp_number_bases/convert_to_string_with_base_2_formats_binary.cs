// vybe-test: csharp/csharp_number_bases/convert_to_string_with_base_2_formats_binary
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Convert.ToString(10,2)).ToString(), "1010");
