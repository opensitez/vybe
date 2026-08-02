// vybe-test: csharp/csharp_nameof_expressions/nameof_record_primary_constructor_parameter
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Book(string Title,int Pages); __Check((nameof(Book.Title)).ToString(), "Title");
