// vybe-test: csharp/csharp_yield_iterators_core/yield_return_enumerable_of_enumerable_flatten_manual
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

System.Collections.Generic.IEnumerable<System.Collections.Generic.IEnumerable<int>> Batches(){yield return new[]{1,2};yield return new[]{3};}
var flat=new System.Collections.Generic.List<int>(); foreach(var batch in Batches()) foreach(var n in batch) flat.Add(n); __P((string.Join(",",flat)).ToString());
__Check("1,2,3");
