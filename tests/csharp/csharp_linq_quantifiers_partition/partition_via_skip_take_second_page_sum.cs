// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_via_skip_take_second_page_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var src=new[]{1,2,3,4,5,6};
__Check((src.Skip(2).Take(2).Sum()).ToString(), "7");
