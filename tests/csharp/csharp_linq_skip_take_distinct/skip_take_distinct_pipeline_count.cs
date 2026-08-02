// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_take_distinct_pipeline_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2,2,3,3,4,5}.Skip(1).Take(5).Distinct();
__Check((r.Count()).ToString(), "3");
