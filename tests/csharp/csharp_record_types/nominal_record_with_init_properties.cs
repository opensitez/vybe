// vybe-test: csharp/csharp_record_types/nominal_record_with_init_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config { public string Host { get; init; } public int Port { get; init; } }
var c = new Config { Host="localhost", Port=8080 };
__Check((c.Host).ToString(), "localhost"); __Check((c.Port).ToString(), "8080");
