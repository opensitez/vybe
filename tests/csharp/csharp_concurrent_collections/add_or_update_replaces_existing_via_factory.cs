// vybe-test: csharp/csharp_concurrent_collections/add_or_update_replaces_existing_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["k"] = 1;
d.AddOrUpdate("k", 0, (key, old) => old + 10);
__Check((d["k"]).ToString(), "11");
