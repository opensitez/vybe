// vybe-test: csharp/csharp_class_indexers/derived_class_indexer_can_call_base_indexer_via_cast
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    protected int[] data = { 1, 2 };
    public virtual int this[int i] { get { return data[i]; } }
}
class Derived : Base {
    public override int this[int i] { get { return base[i] + 10; } }
}
Base item = new Derived();
__Check((item[1]).ToString(), "12");
