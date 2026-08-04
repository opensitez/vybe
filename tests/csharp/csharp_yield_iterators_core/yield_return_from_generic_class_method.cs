// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_generic_class_method
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

class Bag<T>{public System.Collections.Generic.IEnumerable<T> Single(T v){yield return v;}}
__P((new Bag<int>().Single(8).First()).ToString());
__Check("8");
