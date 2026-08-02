// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_generic_class_method
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bag<T>{public System.Collections.Generic.IEnumerable<T> Single(T v){yield return v;}}
__Check((new Bag<int>().Single(8).First()).ToString(), "8");
