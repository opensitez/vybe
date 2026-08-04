// vybe-test: csharp/csharp_with_expression_records/with_nominal_init_port
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

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

record Config{public string Host{get;init;}="localhost"; public int Port{get;init;}=80;} var p=(new Config()) with{Port=443}; __P((p.Port).ToString());
__Check("443");
