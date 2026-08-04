// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_property_setter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

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

class Box {
    int _v;
    public int Value {
        set {
            Report();
            _v = value;
        }
        get => _v;
    }
    void Report([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
var b = new Box(); b.Value = 9; __P((b.Value).ToString());
__Check("Value\n9");
