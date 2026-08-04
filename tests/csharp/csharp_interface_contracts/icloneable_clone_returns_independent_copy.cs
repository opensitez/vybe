// vybe-test: csharp/csharp_interface_contracts/icloneable_clone_returns_independent_copy
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

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

class Box : System.ICloneable {
    public int Value;
    public object Clone() => new Box { Value = Value };
}
var original = new Box { Value=5 };
var copy = (Box)original.Clone();
copy.Value = 99;
__P((original.Value).ToString());
__Check("5");
