// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_custom_class_returned_from_method
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node { public int Value; }
Node Make() { Node n = new(); n.Value = 12; return n; }
__Check((Make().Value).ToString(), "12");
