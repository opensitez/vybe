// vybe-test: csharp/csharp_with_expression_records/with_nested_triple_inner
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Zip(string Code); record Address(string City,Zip Z); record Person(string Name,Address Home); var p=new Person("A",new Address("Oslo",new Zip("01"))); var q=p with{Home=p.Home with{Z=p.Home.Z with{Code="02"}}}; __Check((q.Home.Z.Code).ToString(), "02");
