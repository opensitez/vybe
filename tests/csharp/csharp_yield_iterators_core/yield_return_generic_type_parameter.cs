// vybe-test: csharp/csharp_yield_iterators_core/yield_return_generic_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<T> Echo<T>(T v){yield return v;yield return v;}
__Check((string.Join(",",Echo("x"))).ToString(), "x,x");
