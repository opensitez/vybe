// vybe-test: csharp/csharp_concurrent_collections/get_or_add_returns_existing_value_without_adding
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 5;
__Check((d.GetOrAdd("x", 99)).ToString(), "5");
