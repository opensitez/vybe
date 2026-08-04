// vybe-test: csharp/csharp_virtual_dispatch_semantics/explicit_interface_implementation_not_visible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

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

interface IWorker {
    string Work();
}
class Person : IWorker {
    string IWorker.Work() { return "hidden"; }
    public string Work() { return "public"; }
}
Person person = new Person();
__P((person.Work()).ToString());
__P((((IWorker)person).Work()).ToString());
__Check("public\nhidden");
