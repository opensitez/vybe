// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_default_used_when_not_overridden
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBase<T> where T:IBase<T>{static virtual int Code=>0; static abstract T Build();}
struct S:IBase<S>{public static S Build()=>new S();}
__Check((S.Code).ToString(), "0");
