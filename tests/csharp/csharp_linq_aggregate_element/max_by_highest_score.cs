// vybe-test: csharp/csharp_linq_aggregate_element/max_by_highest_score
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{(N:"a",S:1),(N:"b",S:5),(N:"c",S:3)}.MaxBy(t=>t.S).N).ToString(), "b");
