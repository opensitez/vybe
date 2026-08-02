// vybe-test: csharp/csharp_class_indexers/indexer_on_readonly_wrapper_exposes_underlying_element
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class ReadWrapper {
    readonly int[] data = { 5, 6 };
    public int this[int i] { get { return data[i]; } }
}
__Check((new ReadWrapper()[0]).ToString(), "5");
