// vybe-test: csharp/csharp_class_indexers/derived_class_indexer_can_call_base_indexer_via_cast
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

class Base {
    protected int[] data = { 1, 2 };
    public virtual int this[int i] { get { return data[i]; } }
}
class Derived : Base {
    public override int this[int i] { get { return base[i] + 10; } }
}
Base item = new Derived();
__P((item[1]).ToString());
__Check("12");
