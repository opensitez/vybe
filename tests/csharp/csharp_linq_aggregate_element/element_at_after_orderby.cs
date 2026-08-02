// vybe-test: csharp/csharp_linq_aggregate_element/element_at_after_orderby
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{3,1,2}.OrderBy(x=>x).ElementAt(1)).ToString(), "2");
