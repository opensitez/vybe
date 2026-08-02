// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_custom_class_inferred_from_local
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int Id = 0; }
Widget w = new();
w.Id = 9;
__Check((w.Id).ToString(), "9");
