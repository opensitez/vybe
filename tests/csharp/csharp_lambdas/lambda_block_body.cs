// vybe-test: csharp/csharp_lambdas/lambda_block_body
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var add = (int a, int b) => {
    return a + b;
};
__Check((add(3, 4)).ToString(), "7");
