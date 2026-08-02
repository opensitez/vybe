// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_sum_integers
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4}.Aggregate(0,(acc,x)=>acc+x)).ToString(), "10");
