// vybe-test: csharp/csharp_readonly_members/init_property_settable_only_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{public int Port{get;init;}=80;}
var c=new Config{Port=443};
__Check((c.Port).ToString(), "443");
