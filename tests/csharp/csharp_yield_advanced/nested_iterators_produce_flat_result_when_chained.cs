// vybe-test: csharp/csharp_yield_advanced/nested_iterators_produce_flat_result_when_chained
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Doubles(int n){
    yield return n; yield return n*2;
}
var result=new[]{1,2,3}.SelectMany(Doubles);
__Check((string.Join(",",result)).ToString(), "1,2,2,4,3,6");
