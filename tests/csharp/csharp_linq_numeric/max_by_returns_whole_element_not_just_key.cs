// vybe-test: csharp/csharp_linq_numeric/max_by_returns_whole_element_not_just_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"cat","elephant","ox"};
__Check((words.MaxBy(w=>w.Length)).ToString(), "elephant");
