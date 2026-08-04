// vybe-test: csharp/csharp_init_required_members/init_property_on_positional_record_with_extra_init
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

record User(string Name) { public int Age { get; init; } = 0; }
var u = new User("Bob") { Age = 30 };
__P((u.Name).ToString()); __P((u.Age).ToString());
__Check("Bob\n30");
