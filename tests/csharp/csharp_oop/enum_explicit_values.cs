// vybe-test: csharp/csharp_oop/enum_explicit_values
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum HttpStatus {
    OK = 200,
    NotFound = 404,
    ServerError = 500
}
__Check((HttpStatus.NotFound).ToString(), "404");
