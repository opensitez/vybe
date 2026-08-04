// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_instance_method_on_class
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

class Counter{public System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}}
__P((new Counter().Range(4).Sum()).ToString());
__Check("6");
