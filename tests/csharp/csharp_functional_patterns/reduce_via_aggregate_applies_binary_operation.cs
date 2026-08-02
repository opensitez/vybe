// vybe-test: csharp/csharp_functional_patterns/reduce_via_aggregate_applies_binary_operation
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var product=new[]{1,2,3,4,5}.Aggregate((acc,x)=>acc*x);
__Check((product).ToString(), "120");
