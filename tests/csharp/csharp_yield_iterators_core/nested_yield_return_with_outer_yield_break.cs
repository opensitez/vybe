// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_return_with_outer_yield_break
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

System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in new[]{1,2,3}){if(x==2)yield break;yield return x;}}
__P((string.Join(",",Outer())).ToString());
__Check("1");
