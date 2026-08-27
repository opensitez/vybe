// vybe-test: csharp/interfaces_generics/generic_where_new
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var f = new Factory<Item>();
var item = f.Create();
__P((item.Name).ToString());
__Check("default");

class Factory<T> where T : new() {
    public T Create() { return new T(); }
}

class Item {
    public string Name = "default";
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
