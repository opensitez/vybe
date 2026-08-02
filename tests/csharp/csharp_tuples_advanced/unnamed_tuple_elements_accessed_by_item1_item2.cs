// vybe-test: csharp/csharp_tuples_advanced/unnamed_tuple_elements_accessed_by_item1_item2
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (1, "hello");
__Check((t.Item1).ToString(), "1"); __Check((t.Item2).ToString(), "hello");
