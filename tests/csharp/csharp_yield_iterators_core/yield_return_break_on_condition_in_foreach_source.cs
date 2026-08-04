// vybe-test: csharp/csharp_yield_iterators_core/yield_return_break_on_condition_in_foreach_source
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

System.Collections.Generic.IEnumerable<int> TakeWhilePositive(int[] a){foreach(var n in a){if(n<0)yield break;yield return n;}}
__P((string.Join(",",TakeWhilePositive(new[]{2,4,-1,8}))).ToString());
__Check("2,4");
