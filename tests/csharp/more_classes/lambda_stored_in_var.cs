// vybe-test: csharp/more_classes/lambda_stored_in_var
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var twice = x => x * 2;
        __Check((twice(21)).ToString(), "42");
