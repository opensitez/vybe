// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_calls_abstract_in_default_body
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

interface IWrap<T> where T:IWrap<T>{static abstract T Core(); static virtual T Outer(){return Core();}}
struct Core:IWrap<Core>{public int N; public static Core Core()=>new Core{N=4};}
__P((Core.Outer().N).ToString());
__Check("4");
