// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_empty_arrays
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Array.Empty<int>().SequenceEqual(System.Array.Empty<int>())).ToString(), "True");
