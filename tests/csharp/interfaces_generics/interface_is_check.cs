// vybe-test: csharp/interfaces_generics/interface_is_check
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFlyable { }
class Bird : IFlyable { }
class Fish { }
object b = new Bird();
object f = new Fish();
__Check((b is IFlyable).ToString(), "True");
__Check((f is IFlyable).ToString(), "False");
