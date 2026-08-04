// vybe-test: csharp/csharp_class_indexers/indexer_set_can_update_multiple_fields_atomically
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

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

class PairStore {
    public int First;
    public int Second;
    public int this[int slot] {
        get { return slot == 0 ? First : Second; }
        set { if (slot == 0) First = value; else Second = value; }
    }
}
var store = new PairStore();
store[0] = 3;
store[1] = 9;
__P((store.First).ToString());
__P((store.Second).ToString());
__Check("3\n9");
