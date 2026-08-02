// vybe-test: csharp/csharp_with_expression_records/with_nominal_host_only
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config{public string Host{get;init;} public int Port{get;init;}} var c=new Config{Host="a",Port=1}; var d=c with{Host="b"}; __Check((c.Host).ToString(), "a"); __Check((d.Host).ToString(), "b");
