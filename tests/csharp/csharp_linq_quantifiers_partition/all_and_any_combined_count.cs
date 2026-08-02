// vybe-test: csharp/csharp_linq_quantifiers_partition/all_and_any_combined_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{2,4,6,8};
__Check((data.All(x=>x%2==0)?1:0).ToString(), "1");
__Check((data.Any(x=>x>5)?1:0).ToString(), "1");
