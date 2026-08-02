// vybe-test: csharp/csharp_string_methods/substring_extracts_region_by_start_and_length
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".Substring(6, 5)).ToString(), "world");
