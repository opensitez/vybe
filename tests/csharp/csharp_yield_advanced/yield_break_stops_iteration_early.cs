// vybe-test: csharp/csharp_yield_advanced/yield_break_stops_iteration_early
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

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

System.Collections.Generic.IEnumerable<int> Take(int[] a,int max){
    int count=0;
    foreach(var n in a){
        if(count>=max) yield break;
        yield return n;
        count++;
    }
}
__P((string.Join(",",Take(new[]{1,2,3,4,5},3))).ToString());
__Check("1,2,3");
