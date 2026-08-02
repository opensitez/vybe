// vybe-test: csharp/csharp_linq_complex/aggregate_with_seed_computes_running_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{1,2,3,4,5}.Aggregate(1L,(acc,n)=>acc*n);
__Check((result).ToString(), "120");
