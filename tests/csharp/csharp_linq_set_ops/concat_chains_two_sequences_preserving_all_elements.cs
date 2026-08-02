// vybe-test: csharp/csharp_linq_set_ops/concat_chains_two_sequences_preserving_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result = new[]{1,2}.Concat(new[]{3,4});
__Check((result.Count()).ToString(), "4");
