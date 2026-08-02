// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_no_seed_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{2,3,4}.Aggregate((a,b)=>a*b)).ToString(), "24");
