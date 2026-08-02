// vybe-test: csharp/csharp_abstract_class/abstract_class_holding_state_shared_with_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Counter{
    protected int Count;
    public abstract void Increment();
    public int Value=>Count;
}
class By2:Counter{public override void Increment(){Count+=2;}}
var c=new By2(); c.Increment(); c.Increment();
__Check((c.Value).ToString(), "4");
