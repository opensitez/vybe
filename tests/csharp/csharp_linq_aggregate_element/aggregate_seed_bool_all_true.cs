// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_bool_all_true
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ok=new[]{true,true,true}.Aggregate(true,(acc,x)=>acc&&x);
__Check((ok).ToString(), "True");
