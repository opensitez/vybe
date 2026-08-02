// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_implementors_same_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICode<T> where T:ICode<T>{static abstract int Code();}
struct A:ICode<A>{public static int Code()=>1;} struct B:ICode<B>{public static int Code()=>2;}
__Check((A.Code()+B.Code()).ToString(), "3");
