// vybe-test: csharp/csharp_concurrent_collections/try_remove_extracts_and_deletes_entry
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 7;
__Check((d.TryRemove("x", out int v)).ToString(), "True");
__Check((v).ToString(), "7");
__Check((d.Count).ToString(), "0");
