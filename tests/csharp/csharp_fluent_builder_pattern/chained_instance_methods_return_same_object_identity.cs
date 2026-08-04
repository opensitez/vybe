// vybe-test: csharp/csharp_fluent_builder_pattern/chained_instance_methods_return_same_object_identity
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

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
__P((same.Id() == builder.Id() ? "Y" : "N").ToString());
__P((builder.Build()).ToString());
__Check("Y\n5");
