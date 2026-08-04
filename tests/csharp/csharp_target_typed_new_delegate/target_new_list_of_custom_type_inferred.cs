// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_of_custom_type_inferred
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

class Item { public string Name = ""; }
System.Collections.Generic.List<Item> items = new();
items.Add(new Item { Name = "tool" });
__P((items[0].Name).ToString());
__Check("tool");
