// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_implementors_same_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface ICode<T> where T:ICode<T>{static abstract int Code();}
struct A:ICode<A>{public static int Code()=>1;} struct B:ICode<B>{public static int Code()=>2;}
__P((A.Code()+B.Code()).ToString());
__Check("3");
