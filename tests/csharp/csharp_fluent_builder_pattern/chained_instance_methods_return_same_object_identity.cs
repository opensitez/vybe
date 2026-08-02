// vybe-test: csharp/csharp_fluent_builder_pattern/chained_instance_methods_return_same_object_identity
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Builder {
    static int nextId;
    int id = ++nextId;
    int total;
    public Builder Add(int value) { total += value; return this; }
    public int Id() { return id; }
    public int Build() { return total; }
}
var builder = new Builder();
var same = builder.Add(2).Add(3);
__Check((same.Id() == builder.Id() ? "Y" : "N").ToString(), "Y");
__Check((builder.Build()).ToString(), "5");
