// vybe-test: csharp/csharp_generics_where/where_new_constraint_allows_parameterless_construction
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

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

T Build<T>() where T:new()=>new T();
class Box{public int V=7;}
__P((Build<Box>().V).ToString());
__Check("7");
