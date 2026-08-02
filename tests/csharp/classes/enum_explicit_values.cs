// vybe-test: csharp/classes/enum_explicit_values
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Status { Ok = 200, NotFound = 404, Error = 500 }
        __Check((Status.Ok).ToString(), "200");
        __Check((Status.NotFound).ToString(), "404");
