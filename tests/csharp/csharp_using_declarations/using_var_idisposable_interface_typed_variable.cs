// vybe-test: csharp/csharp_using_declarations/using_var_idisposable_interface_typed_variable
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "use");}}
using System.IDisposable x=new R("iface"); __Check(("use").ToString(), "iface");
