// vybe-test: csharp/advanced/lambda_expression_arrow
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var fn = x => x + 1;
        __Check((fn(9)).ToString(), "10");
