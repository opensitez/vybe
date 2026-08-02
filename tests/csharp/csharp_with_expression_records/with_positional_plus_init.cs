// vybe-test: csharp/csharp_with_expression_records/with_positional_plus_init
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name){public int Age{get;init;}} var v=(new User("Ada"){Age=20}) with{Age=21}; __Check((v.Age).ToString(), "21");
