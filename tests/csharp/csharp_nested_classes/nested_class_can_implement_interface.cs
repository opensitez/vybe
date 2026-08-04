// vybe-test: csharp/csharp_nested_classes/nested_class_can_implement_interface
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

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

interface IValue{int Get();}
class Host{public class Impl:IValue{public int Get()=>5;}}
IValue v=new Host.Impl();
__P((v.Get()).ToString());
__Check("5");
