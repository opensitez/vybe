// vybe-test: csharp/csharp_abstract_class/abstract_class_holding_state_shared_with_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

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

abstract class Counter{
    protected int Count;
    public abstract void Increment();
    public int Value=>Count;
}
class By2:Counter{public override void Increment(){Count+=2;}}
var c=new By2(); c.Increment(); c.Increment();
__P((c.Value).ToString());
__Check("4");
