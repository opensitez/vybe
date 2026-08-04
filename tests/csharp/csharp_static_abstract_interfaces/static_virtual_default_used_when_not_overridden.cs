// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_default_used_when_not_overridden
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

interface IBase<T> where T:IBase<T>{static virtual int Code=>0; static abstract T Build();}
struct S:IBase<S>{public static S Build()=>new S();}
__P((S.Code).ToString());
__Check("0");
