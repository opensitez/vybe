// vybe-test: csharp/csharp_yield_iterators_core/yield_return_enumerable_as_return_type_of_helper
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Build(){yield return 3;yield return 5;}
int Total(){return Build().Sum();} __Check((Total()).ToString(), "8");
