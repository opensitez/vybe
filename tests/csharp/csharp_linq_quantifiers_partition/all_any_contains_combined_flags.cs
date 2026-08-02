// vybe-test: csharp/csharp_linq_quantifiers_partition/all_any_contains_combined_flags
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var xs=new[]{1,2,3};
__Check((xs.All(x=>x>0)?1:0).ToString(), "1");
__Check((xs.Any(x=>x==2)?1:0).ToString(), "1");
__Check((xs.Contains(4)?1:0).ToString(), "0");
