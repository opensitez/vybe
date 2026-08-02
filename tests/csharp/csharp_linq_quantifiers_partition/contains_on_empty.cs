// vybe-test: csharp/csharp_linq_quantifiers_partition/contains_on_empty
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Array.Empty<int>().Contains(1)).ToString(), "False");
