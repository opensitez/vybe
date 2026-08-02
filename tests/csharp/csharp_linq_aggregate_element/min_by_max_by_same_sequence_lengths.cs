// vybe-test: csharp/csharp_linq_aggregate_element/min_by_max_by_same_sequence_lengths
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"go","stop","run"};
__Check((words.MinBy(w=>w.Length).Length).ToString(), "2");
__Check((words.MaxBy(w=>w.Length).Length).ToString(), "4");
