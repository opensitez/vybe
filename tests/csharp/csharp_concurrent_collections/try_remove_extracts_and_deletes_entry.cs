// vybe-test: csharp/csharp_concurrent_collections/try_remove_extracts_and_deletes_entry
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 7;
__P((d.TryRemove("x", out int v)).ToString());
__P((v).ToString());
__P((d.Count).ToString());
__Check("True\n7\n0");
