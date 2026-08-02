// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_action_from_void_method_group
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 0;
void Bump() { total++; }
System.Action bump = Bump;
bump(); bump();
__Check((total).ToString(), "2");
