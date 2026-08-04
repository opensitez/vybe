// vybe-test: csharp/csharp_init_required_members/init_property_mixed_with_settable_property
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

class Item { public int Id { get; init; } public string Label { get; set; } = ""; }
var i = new Item { Id = 7 };
i.Label = "tool";
__P((i.Id).ToString()); __P((i.Label).ToString());
__Check("7\ntool");
