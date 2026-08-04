// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_local_function
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

System.Collections.Generic.IEnumerable<int> Outer(){System.Collections.Generic.IEnumerable<int> Inner(){yield return 3;} foreach(var x in Inner())yield return x;}
__P((Outer().First()).ToString());
__Check("3");
