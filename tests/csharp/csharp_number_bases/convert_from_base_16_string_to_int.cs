// vybe-test: csharp/csharp_number_bases/convert_from_base_16_string_to_int
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Convert.ToInt32("ff",16)).ToString(), "255");
