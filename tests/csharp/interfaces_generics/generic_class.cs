// vybe-test: csharp/interfaces_generics/generic_class
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var intBox = new Box<int>(42);
var strBox = new Box<string>("hello");
__P((intBox.Value).ToString());
__P((strBox.Value).ToString());
__Check("42\nhello");

class Box<T> {
    public T Value;
    public Box(T val) { Value = val; }
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
