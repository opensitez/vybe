// vybe-test: csharp/csharp_interface_contracts/iequatable_equals_compared_by_sorted_set_deduplication
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Id : System.IEquatable<Id> {
    public int Value;
    public bool Equals(Id other) => other?.Value == Value;
    public override bool Equals(object o) => o is Id i && Equals(i);
    public override int GetHashCode() => Value;
}
var set = new System.Collections.Generic.HashSet<Id>(
    System.Collections.Generic.EqualityComparer<Id>.Default);
set.Add(new Id{Value=1}); set.Add(new Id{Value=1});
__Check((set.Count).ToString(), "1");
