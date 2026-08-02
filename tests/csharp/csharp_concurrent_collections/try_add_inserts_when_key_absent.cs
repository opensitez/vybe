// vybe-test: csharp/csharp_concurrent_collections/try_add_inserts_when_key_absent
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
__Check((d.TryAdd("a", 1)).ToString(), "True");
__Check((d["a"]).ToString(), "1");
