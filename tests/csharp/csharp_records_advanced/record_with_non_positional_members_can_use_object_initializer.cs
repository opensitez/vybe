// vybe-test: csharp/csharp_records_advanced/record_with_non_positional_members_can_use_object_initializer
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

record Theme { public string Name { get; init; } public int Version { get; init; } } var theme = new Theme { Name = "light", Version = 2 }; __P((theme.Name + ":" + theme.Version).ToString());
__Check("light:2");
