// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_min_running
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var min=new[]{3,1,4,1,5}.Aggregate(int.MaxValue,(acc,x)=>x<acc?x:acc);
__Check((min).ToString(), "1");
