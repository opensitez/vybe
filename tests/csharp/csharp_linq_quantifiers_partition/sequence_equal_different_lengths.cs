// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_different_lengths
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.SequenceEqual(new[]{1,2})).ToString(), "False");
