// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_then_element_at
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var running=new[]{1,2,3}.Aggregate(new int[]{0},(acc,x)=>new int[]{acc[0]+x});
__Check((running[0]).ToString(), "6");
