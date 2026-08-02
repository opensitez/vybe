// vybe-test: csharp/csharp_linq_set_ops/sequence_equal_returns_false_for_different_order
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.SequenceEqual(new[]{3,2,1})).ToString(), "False");
