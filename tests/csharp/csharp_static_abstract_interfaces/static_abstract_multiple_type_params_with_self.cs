// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_type_params_with_self
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

interface IPair<TSelf,TVal> where TSelf:IPair<TSelf,TVal>{static abstract TSelf Of(TVal v);}
struct Holder:IPair<Holder,int>{public int Data; public static Holder Of(int v)=>new Holder{Data=v};}
__P((Holder.Of(6).Data).ToString());
__Check("6");
