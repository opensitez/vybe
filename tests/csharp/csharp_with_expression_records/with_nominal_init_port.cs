// vybe-test: csharp/csharp_with_expression_records/with_nominal_init_port
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config{public string Host{get;init;}="localhost"; public int Port{get;init;}=80;} var p=(new Config()) with{Port=443}; __Check((p.Port).ToString(), "443");
