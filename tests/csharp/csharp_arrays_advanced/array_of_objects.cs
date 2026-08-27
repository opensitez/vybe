// vybe-test: csharp/csharp_arrays_advanced/array_of_objects
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var items = new[] { new Item("a"), new Item("b"), new Item("c") }
;
foreach (var item in items) {
    __P((item.Name).ToString());
}
__Check("a\nb\nc");

class Item {
    public string Name;
    public Item(string n) { Name = n; }
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
