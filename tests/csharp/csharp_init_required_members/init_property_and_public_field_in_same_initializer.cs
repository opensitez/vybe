// vybe-test: csharp/csharp_init_required_members/init_property_and_public_field_in_same_initializer
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

class Form { public string Title { get; init; } public int Version; }
var f = new Form { Title = "main", Version = 2 };
__P((f.Title).ToString()); __P((f.Version).ToString());
__Check("main\n2");
