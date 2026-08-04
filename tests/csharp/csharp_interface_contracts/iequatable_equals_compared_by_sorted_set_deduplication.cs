// vybe-test: csharp/csharp_interface_contracts/iequatable_equals_compared_by_sorted_set_deduplication
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

class Id : System.IEquatable<Id> {
    public int Value;
    public bool Equals(Id other) => other?.Value == Value;
    public override bool Equals(object o) => o is Id i && Equals(i);
    public override int GetHashCode() => Value;
}
var set = new System.Collections.Generic.HashSet<Id>(
    System.Collections.Generic.EqualityComparer<Id>.Default);
set.Add(new Id{Value=1}); set.Add(new Id{Value=1});
__P((set.Count).ToString());
__Check("1");
