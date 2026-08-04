// vybe-test: csharp/csharp_init_required_members/with_expression_changes_init_property_on_record
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

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

record Config { public int Port { get; init; } = 80; }
var a = new Config();
var b = a with { Port = 9000 };
__P((a.Port).ToString()); __P((b.Port).ToString());
__Check("80\n9000");
