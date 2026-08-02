// vybe-test: csharp/csharp_with_expression/with_expression_on_record_with_init_property
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config { public string Host { get; init; } public int Port { get; init; } }
var base_ = new Config { Host = "localhost", Port = 80 };
var prod = base_ with { Port = 443 };
__Check((prod.Host).ToString(), "localhost");
__Check((prod.Port).ToString(), "443");
