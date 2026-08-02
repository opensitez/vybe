// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_with_index_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sum=new[]{10,20,30}.Aggregate(0,(acc,x,i)=>acc+x+i);
__Check((sum).ToString(), "63");
