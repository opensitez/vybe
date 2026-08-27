// vybe-test: csharp/csharp_class_indexers/indexer_set_can_update_multiple_fields_atomically
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

var store = new PairStore();
store[0] = 3;
store[1] = 9;
__P((store.First).ToString());
__P((store.Second).ToString());
__Check("3\n9");

class PairStore {
    public int First;
    public int Second;
    public int this[int slot] {
        get { return slot == 0 ? First : Second; }
        set { if (slot == 0) First = value; else Second = value; }
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
