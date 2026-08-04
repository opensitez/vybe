// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_custom_class_as_field_initializer
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

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

class Holder { public System.Collections.Generic.List<int> items = new(); }
var h = new Holder();
h.items.Add(6);
__P((h.items[0]).ToString());
__Check("6");
