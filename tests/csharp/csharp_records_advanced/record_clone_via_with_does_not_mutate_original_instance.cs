// vybe-test: csharp/csharp_records_advanced/record_clone_via_with_does_not_mutate_original_instance
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

record User(string Name, int Age); var before = new User("Ada", 30); var after = before with { Name = "Grace" }; __P((before.Name).ToString()); __P((after.Name).ToString());
__Check("Ada\nGrace");
