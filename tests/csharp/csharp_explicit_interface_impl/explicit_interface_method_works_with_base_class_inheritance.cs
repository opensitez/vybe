// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_works_with_base_class_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

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

interface ILabel { string Label(); }
class BaseItem {
    protected string prefix = "base";
}
class TaggedItem : BaseItem, ILabel {
    string ILabel.Label() { return prefix + "/tag"; }
}
__P((((ILabel)new TaggedItem()).Label()).ToString());
__Check("base/tag");
