// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_no_seed_max_via_fold
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{3,1,4}.Aggregate((a,b)=>a>b?a:b)).ToString(), "4");
