// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_nominal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Config{public int Port{get;init;}=80;} var c=new Config{Port=8080}; var d=c with{Port=443}; __Check((c.Port).ToString(), "8080"); __Check((d.Port).ToString(), "443");
