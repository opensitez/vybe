// vybe-test: csharp/linq_lambdas/lambda_passed_to_method
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Processor {
    public int Apply(int value, Func<int, int> transform) {
        return transform(value);
    }
}
var p = new Processor();
__Check((p.Apply(5, x => x * x)).ToString(), "25");
__Check((p.Apply(5, x => x + 10)).ToString(), "15");
