// vybe-test: csharp/csharp_records_advanced/with_expression_copies_record_and_changes_one_member
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

record User(string Name, int Age); var user = new User("Ada", 20); var updated = user with { Age = 21 }; __P((user.Age).ToString()); __P((updated.Age).ToString());
__Check("20\n21");
