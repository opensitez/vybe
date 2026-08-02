// vybe-test: csharp/csharp_new_features/target_typed_new_in_constructor_argument
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public System.Collections.Generic.List<int> Items; public Box(System.Collections.Generic.List<int> i){Items=i;} }
var b = new Box(new());
b.Items.Add(9);
__Check((b.Items.Count).ToString(), "1");
