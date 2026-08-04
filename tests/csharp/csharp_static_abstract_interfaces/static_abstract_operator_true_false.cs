// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_operator_true_false
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

interface ITest<T> where T:ITest<T>{static abstract bool operator true(T v); static abstract bool operator false(T v);}
struct Flag:ITest<Flag>{public bool On; public static bool operator true(Flag v)=>v.On; public static bool operator false(Flag v)=>!v.On;}
__P((new Flag{On=true}?1:0).ToString());
__Check("1");
