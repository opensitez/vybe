// vybe-test: csharp/csharp_with_expression/with_expression_on_record_with_init_property
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

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
var base_ = new Config { Host = "localhost", Port = 80 };
var prod = base_ with { Port = 443 };
__P((prod.Host).ToString());
__P((prod.Port).ToString());
__Check("localhost\n443");
