// vybe-test: csharp/csharp_class_indexers/indexer_set_can_update_multiple_fields_atomically
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((store.First).ToString(), "3");
__Check((store.Second).ToString(), "9");
