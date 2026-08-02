// vybe-test: csharp/csharp_object_initializers/with_expression_on_nominal_record_is_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config{public int Port{get;init;}=80;}
var cfg=new Config() with{Port=443};
__Check((cfg.Port).ToString(), "443");
