// vybe-test: csharp/csharp_concurrent_collections/try_add_returns_false_when_key_present
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d.TryAdd("a", 1);
__Check((d.TryAdd("a", 9)).ToString(), "False");
