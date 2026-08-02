// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_action_inferred_from_lambda
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 0;
System.Action tick = () => count++;
tick(); tick();
__Check((count).ToString(), "2");
