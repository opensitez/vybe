// vybe-test: csharp/csharp_with_expression_records/with_nominal_chain
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record C{public int A{get;init;} public int B{get;init;}} var e=((new C{A=1,B=2}) with{A=3}) with{B=4}; __Check((e.A).ToString(), "3"); __Check((e.B).ToString(), "4");
