// vybe-test: csharp/csharp_class_indexers/indexer_get_can_compute_derived_value_from_state
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale {
    int factor = 2;
    public int this[int input] { get { return input * factor; } }
}
__Check((new Scale()[5]).ToString(), "10");
