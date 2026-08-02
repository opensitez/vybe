// vybe-test: csharp/csharp_virtual_dispatch_semantics/chained_base_constructor_initializes_before_derived_fields
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    protected string token;
    public Base(string token) { this.token = token; }
}
class Child : Base {
    public string Label;
    public Child(string token, string label) : base(token) { Label = label; }
    public string Read() { return token + ":" + Label; }
}
__Check((new Child("id", "name").Read()).ToString(), "id:name");
