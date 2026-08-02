// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_same_elements_different_order
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.SequenceEqual(new[]{3,2,1})).ToString(), "False");
