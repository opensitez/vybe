// vybe-test: csharp/csharp_interface_contracts/icloneable_clone_returns_independent_copy
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((original.Value).ToString(), "5");
