// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_override_changes_code
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBase<T> where T:IBase<T>{static virtual int Code=>0; static abstract T Build();}
struct S:IBase<S>{public static S Build()=>new S(); public static int Code=>5;}
__Check((S.Code).ToString(), "5");
