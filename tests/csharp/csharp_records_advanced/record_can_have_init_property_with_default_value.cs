// vybe-test: csharp/csharp_records_advanced/record_can_have_init_property_with_default_value
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

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

record User { public string Name { get; init; } = "guest"; } __P((new User().Name).ToString());
__Check("guest");
