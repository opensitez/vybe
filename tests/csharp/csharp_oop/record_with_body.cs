// vybe-test: csharp/csharp_oop/record_with_body
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var p = new Product("Widget", 9.99);
__P((p.Display()).ToString());
__Check("Widget: $9.99");

record Product(string Name, double Price) {
    public string Display() { return Name + ": $" + Price; }
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
