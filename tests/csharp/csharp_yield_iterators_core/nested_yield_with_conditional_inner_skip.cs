// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_with_conditional_inner_skip
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

System.Collections.Generic.IEnumerable<int> Inner(bool ok){if(ok)yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(bool ok){foreach(var x in Inner(ok))yield return x;yield return 1;}
__P((string.Join(",",Outer(false))).ToString());
__Check("1");
