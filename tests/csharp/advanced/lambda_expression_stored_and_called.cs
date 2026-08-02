// vybe-test: csharp/advanced/lambda_expression_stored_and_called
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var twice = x => x * 2;
        var result = twice(5);
        __Check((result).ToString(), "10");
