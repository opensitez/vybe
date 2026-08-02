// vybe-test: csharp/csharp_linq_quantifiers_partition/all_strings_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{"a","b"}.All(s=>s.Length>0)).ToString(), "True");
