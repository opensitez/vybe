// vybe-test: csharp/csharp_functional_patterns/map_then_filter_then_reduce_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{1,2,3,4,5}
    .Select(x=>x*x)
    .Where(x=>x>5)
    .Sum();
__Check((result).ToString(), "50");
