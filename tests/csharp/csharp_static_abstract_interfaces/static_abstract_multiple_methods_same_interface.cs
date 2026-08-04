// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_methods_same_interface
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

interface IBoth<T> where T:IBoth<T>{static abstract T FromInt(int n); static abstract T FromString(string s);}
struct Dual:IBoth<Dual>{public string Text; public static Dual FromInt(int n)=>new Dual{Text=n.ToString()}; public static Dual FromString(string s)=>new Dual{Text=s};}
__P((Dual.FromString("ok").Text).ToString());
__Check("ok");
