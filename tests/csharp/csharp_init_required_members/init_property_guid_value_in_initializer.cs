// vybe-test: csharp/csharp_init_required_members/init_property_guid_value_in_initializer
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

class Ref { public System.Guid Id { get; init; } }
var id = new System.Guid("11111111-1111-1111-1111-111111111111");
var r = new Ref { Id = id };
__P((r.Id == id).ToString());
__Check("True");
