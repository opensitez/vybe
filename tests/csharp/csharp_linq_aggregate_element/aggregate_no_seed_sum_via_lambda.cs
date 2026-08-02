// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_no_seed_sum_via_lambda
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3}.Aggregate((a,b)=>a+b)).ToString(), "6");
