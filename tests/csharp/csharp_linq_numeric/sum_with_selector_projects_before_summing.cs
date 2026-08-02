// vybe-test: csharp/csharp_linq_numeric/sum_with_selector_projects_before_summing
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"hello","world","foo"};
__Check((words.Sum(w=>w.Length)).ToString(), "13");
