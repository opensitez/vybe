// vybe-test: csharp/csharp_yield_iterators_core/yield_return_select_where_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

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

System.Collections.Generic.IEnumerable<int> N(){for(int i=0;i<6;i++)yield return i;}
__P((N().Where(x=>x%2==0).Select(x=>x*10).Sum()).ToString());
__Check("60");
