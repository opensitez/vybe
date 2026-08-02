// vybe-test: csharp/csharp_linq_aggregate_element/element_at_last_index
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{10,20,30}.ElementAt(2)).ToString(), "30");
