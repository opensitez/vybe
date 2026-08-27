// vybe-test: csharp/csharp_virtual_dispatch_semantics/explicit_interface_implementation_not_visible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Person person = new Person();
__P((person.Work()).ToString());
__P((((IWorker)person).Work()).ToString());
__Check("public\nhidden");

interface IWorker {
    string Work();
}

class Person : IWorker {
    string IWorker.Work() { return "hidden"; }
    public string Work() { return "public"; }
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
