// vybe-test: csharp/interfaces_generics/interface_is_check
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

interface IFlyable { }
class Bird : IFlyable { }
class Fish { }
object b = new Bird();
object f = new Fish();
__P((b is IFlyable).ToString());
__P((f is IFlyable).ToString());
__Check("True\nFalse");
