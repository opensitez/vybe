// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_is_called_via_interface_reference
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreeter { string Speak(); }
class Person : IGreeter {
    string IGreeter.Speak() { return "hello"; }
}
IGreeter greeter = new Person();
__Check((greeter.Speak()).ToString(), "hello");
