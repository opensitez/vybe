// vybe-test: csharp/csharp_linq_advanced/element_at_returns_item_at_given_index
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{10,20,30}.ElementAt(1)).ToString(), "20");
