// vybe-test: csharp/csharp_functional_patterns/method_chaining_builds_fluent_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{5,3,8,1,4}
    .Where(x=>x>2)
    .OrderBy(x=>x)
    .Select(x=>x*10)
    .First();
__Check((result).ToString(), "30");
