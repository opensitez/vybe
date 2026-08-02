// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_works_with_base_class_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((((ILabel)new TaggedItem()).Label()).ToString(), "base/tag");
