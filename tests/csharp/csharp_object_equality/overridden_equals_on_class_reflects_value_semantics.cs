// vybe-test: csharp/csharp_object_equality/overridden_equals_on_class_reflects_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Money {
    public int Amount;
    public override bool Equals(object obj) =>
        obj is Money m && m.Amount == Amount;
    public override int GetHashCode() => Amount;
}
var x = new Money { Amount = 5 };
var y = new Money { Amount = 5 };
__Check((x.Equals(y)).ToString(), "True");
__Check((object.ReferenceEquals(x, y)).ToString(), "False");
