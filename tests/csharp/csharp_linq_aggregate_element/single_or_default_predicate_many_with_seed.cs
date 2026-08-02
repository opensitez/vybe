// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_predicate_many_with_seed
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{2,2,3}.SingleOrDefault(77,x=>x==2)).ToString(), "77");
