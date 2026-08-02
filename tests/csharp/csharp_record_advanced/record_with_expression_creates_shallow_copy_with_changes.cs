// vybe-test: csharp/csharp_record_advanced/record_with_expression_creates_shallow_copy_with_changes
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config(int Port,string Host);
var c1=new Config(80,"localhost");
var c2=c1 with{Port=443};
__Check((c1.Port).ToString(), "80"); __Check((c2.Port).ToString(), "443");
__Check((c2.Host).ToString(), "localhost");
