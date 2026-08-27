// vybe-test: csharp/csharp_concurrent_dictionary_atomic_ops/concurrent_dict_atomic_case_16

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var cd = new System.Collections.Concurrent.ConcurrentDictionary<string, int>();
bool a1 = cd.TryAdd("k_16", 16);
bool a2 = cd.TryAdd("k_16", 32);
bool up = cd.TryUpdate("k_16", 160, 16);
__P(a1.ToString());
__P(a2.ToString());
__P(up.ToString());
__P(cd["k_16"].ToString());
__Check("True\nFalse\nTrue\n160");
