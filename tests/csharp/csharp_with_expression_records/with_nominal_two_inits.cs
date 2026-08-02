// vybe-test: csharp/csharp_with_expression_records/with_nominal_two_inits
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Theme{public string Name{get;init;} public int Ver{get;init;}} var u=(new Theme{Name="dark",Ver=1}) with{Name="light",Ver=2}; __Check((u.Name).ToString(), "light"); __Check((u.Ver).ToString(), "2");
