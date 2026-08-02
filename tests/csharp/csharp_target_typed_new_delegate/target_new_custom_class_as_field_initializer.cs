// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_custom_class_as_field_initializer
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder { public System.Collections.Generic.List<int> items = new(); }
var h = new Holder();
h.items.Add(6);
__Check((h.items[0]).ToString(), "6");
