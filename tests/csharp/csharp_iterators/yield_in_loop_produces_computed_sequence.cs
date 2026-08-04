// vybe-test: csharp/csharp_iterators/yield_in_loop_produces_computed_sequence
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

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

System.Collections.Generic.IEnumerable<int> Range(int n) {
    for(int i=0; i<n; i++) yield return i;
}
int sum=0;
foreach(var x in Range(5)) sum+=x;
__P((sum).ToString());
__Check("10");
