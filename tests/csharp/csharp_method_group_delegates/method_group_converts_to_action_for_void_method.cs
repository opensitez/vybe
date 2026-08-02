// vybe-test: csharp/csharp_method_group_delegates/method_group_converts_to_action_for_void_method
// origin: languages/csharp/tests/csharp/test_csharp_method_group_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 0;
void Bump() { total++; }
System.Action bump = Bump;
bump();
__Check((total).ToString(), "1");
