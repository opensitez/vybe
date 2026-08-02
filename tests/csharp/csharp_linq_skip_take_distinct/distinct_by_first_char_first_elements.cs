// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_first_char_first_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"cat","car","dog","dot"}.DistinctBy(s=>s[0]);
__Check((r.First()).ToString(), "cat"); __Check((r.Last()).ToString(), "dog");
