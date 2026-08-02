// vybe-test: csharp/csharp_linq_advanced/min_by_returns_element_with_minimum_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"a","bbb","cc"};
__Check((words.MinBy(w=>w.Length)).ToString(), "a");
