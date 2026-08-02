// vybe-test: csharp/csharp_linq_quantifiers_partition/any_after_where_no_matches
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,3,5}.Where(x=>x%2==0).Any()).ToString(), "False");
