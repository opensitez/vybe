// vybe-test: csharp/csharp_linq_quantifiers_partition/contains_and_any_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{1,2,3,4};
__Check((data.Contains(3)?1:0).ToString(), "1");
__Check((data.Any(x=>x>3)?1:0).ToString(), "1");
