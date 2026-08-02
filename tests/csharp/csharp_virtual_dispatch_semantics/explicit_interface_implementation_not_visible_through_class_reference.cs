// vybe-test: csharp/csharp_virtual_dispatch_semantics/explicit_interface_implementation_not_visible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((person.Work()).ToString(), "public");
__Check((((IWorker)person).Work()).ToString(), "hidden");
