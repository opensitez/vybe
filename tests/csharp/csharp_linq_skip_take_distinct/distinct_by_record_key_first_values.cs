// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_record_key_first_values
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{(K:1,V:"a"),(K:1,V:"b"),(K:2,V:"c")}.DistinctBy(t=>t.K);
__Check((r.First().V).ToString(), "a"); __Check((r.Last().V).ToString(), "c");
