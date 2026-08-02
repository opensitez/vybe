// vybe-test: csharp/csharp_linq_numeric/min_with_custom_selector
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"cat","elephant","ox"};
__Check((words.Min(w=>w.Length)).ToString(), "2");
