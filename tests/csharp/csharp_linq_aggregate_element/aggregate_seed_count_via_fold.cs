// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_count_via_fold
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var count=new[]{1,2,3}.Aggregate(0,(acc,x)=>acc+1);
__Check((count).ToString(), "3");
