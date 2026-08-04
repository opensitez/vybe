// vybe-test: csharp/csharp_yield_iterators_core/yield_break_in_loop_stops_at_limit
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

System.Collections.Generic.IEnumerable<int> Take(int max){for(int i=0;i<10;i++){if(i>=max)yield break;yield return i;}}
__P((string.Join(",",Take(3))).ToString());
__Check("0,1,2");
