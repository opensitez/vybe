// vybe-test: csharp/csharp_linq_quantifiers_partition/contains_after_distinct
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,1,2}.Distinct().Contains(2)).ToString(), "True");
