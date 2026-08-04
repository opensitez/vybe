// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_local_state_accumulator
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

System.Collections.Generic.IEnumerable<int> Running(){int s=0; for(int i=1;i<=3;i++){s+=i;yield return s;}}
__P((string.Join(",",Running())).ToString());
__Check("1,3,6");
