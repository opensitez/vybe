// vybe-test: csharp/csharp_object_equality/equal_hash_codes_required_for_equal_objects
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

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

class Key {
    public int Id;
    public override bool Equals(object obj) => obj is Key k && k.Id == Id;
    public override int GetHashCode() => Id;
}
var x = new Key { Id = 7 };
var y = new Key { Id = 7 };
__P((x.GetHashCode() == y.GetHashCode()).ToString());
__Check("True");
