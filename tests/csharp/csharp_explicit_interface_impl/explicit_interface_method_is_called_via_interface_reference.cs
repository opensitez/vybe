// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_is_called_via_interface_reference
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

IGreeter greeter = new Person();
__P((greeter.Speak()).ToString());
__Check("hello");

interface IGreeter { string Speak(); }

class Person : IGreeter {
    string IGreeter.Speak() { return "hello"; }
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
