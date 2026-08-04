// vybe-test: csharp/csharp_reflection_emit/type_get_interfaces_includes_implemented_interfaces
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

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

interface IFoo{}
class Foo:IFoo{}
bool has=System.Array.Exists(typeof(Foo).GetInterfaces(),t=>t==typeof(IFoo));
__P((has).ToString());
__Check("True");
