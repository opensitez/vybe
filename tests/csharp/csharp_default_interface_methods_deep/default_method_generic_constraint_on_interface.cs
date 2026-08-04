// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_generic_constraint_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

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

interface ICompare<T> where T:System.IComparable<T>{int Cmp(T a,T b)=>a.CompareTo(b);} class S:ICompare<int>{} __P((new S().Cmp(3,7)).ToString());
__Check("-1");
