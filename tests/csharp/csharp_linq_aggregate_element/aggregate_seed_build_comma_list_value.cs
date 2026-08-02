// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_build_comma_list_value
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text=new[]{1,2,3}.Aggregate("",(acc,x)=>acc==""?x.ToString():acc+","+x);
__Check((text).ToString(), "1,2,3");
