// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_yields_concatenated_strings
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

System.Collections.Generic.IEnumerable<string> Words(){yield return "a";yield return "b";}
System.Collections.Generic.IEnumerable<string> Twice(){foreach(var w in Words())yield return w;foreach(var w in Words())yield return w;}
__P((string.Join("",Twice())).ToString());
__Check("a,b,a,b");
