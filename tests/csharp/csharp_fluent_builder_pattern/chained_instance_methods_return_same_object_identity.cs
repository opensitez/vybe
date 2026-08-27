// vybe-test: csharp/csharp_fluent_builder_pattern/chained_instance_methods_return_same_object_identity
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

using static __Harness;

var builder = new Builder();
var same = builder.Add(2).Add(3);
__P((same.Id() == builder.Id() ? "Y" : "N").ToString());
__P((builder.Build()).ToString());
__Check("Y\n5");

class Builder {
    static int nextId;
    int id = ++nextId;
    int total;
    public Builder Add(int value) { total += value; return this; }
    public int Id() { return id; }
    public int Build() { return total; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
