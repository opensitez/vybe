// vybe-test: csharp/csharp_reflection_emit/type_get_interfaces_includes_implemented_interfaces
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFoo{}
class Foo:IFoo{}
bool has=System.Array.Exists(typeof(Foo).GetInterfaces(),t=>t==typeof(IFoo));
__Check((has).ToString(), "True");
