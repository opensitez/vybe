// vybe-test: csharp/csharp_record_types/nominal_record_with_init_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

record Config { public string Host { get; init; } public int Port { get; init; } }
var c = new Config { Host="localhost", Port=8080 };
__P((c.Host).ToString()); __P((c.Port).ToString());
__Check("localhost\n8080");
