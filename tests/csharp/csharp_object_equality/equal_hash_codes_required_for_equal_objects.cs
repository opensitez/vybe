// vybe-test: csharp/csharp_object_equality/equal_hash_codes_required_for_equal_objects
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((x.GetHashCode() == y.GetHashCode()).ToString(), "True");
