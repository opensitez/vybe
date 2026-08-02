// vybe-test: csharp/csharp_lambdas/lambda_closure
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int counter = 0;
var inc = () => { counter++; };
inc();
inc();
inc();
__Check((counter).ToString(), "3");
