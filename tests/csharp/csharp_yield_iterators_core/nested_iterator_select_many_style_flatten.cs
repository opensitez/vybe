// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_select_many_style_flatten
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

System.Collections.Generic.IEnumerable<int> Pair(int n){yield return n;yield return n+1;}
__P((string.Join(",",new[]{1,2}.SelectMany(Pair))).ToString());
__Check("1,2,2,3");
