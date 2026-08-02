// vybe-test: csharp/csharp_linq_aggregate_element/single_with_predicate_one_match
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4}.Single(x=>x==3)).ToString(), "3");
