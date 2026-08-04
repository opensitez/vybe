// vybe-test: csharp/csharp_concurrent_collections/get_or_add_inserts_when_key_absent
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
__P((d.GetOrAdd("new", 42)).ToString());
__Check("42");
