// vybe-test: csharp/csharp_new_features/nameof_on_type_member_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int Count; }
__Check((nameof(Widget.Count)).ToString(), "Count");
