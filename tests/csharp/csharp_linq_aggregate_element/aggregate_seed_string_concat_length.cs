// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_string_concat_length
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s=new[]{"a","b","c"}.Aggregate("",(acc,x)=>acc+x);
__Check((s.Length).ToString(), "3");
