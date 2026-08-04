// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_explicit_ienumerable_interface
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

class Nums:System.Collections.Generic.IEnumerable<int>{public System.Collections.Generic.IEnumerator<int> GetEnumerator(){yield return 2;yield return 4;}System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();}
__P((new Nums().Sum()).ToString());
__Check("6");
