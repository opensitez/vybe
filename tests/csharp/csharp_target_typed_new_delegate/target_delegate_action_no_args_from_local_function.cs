// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_action_no_args_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 0;
void Reset() { n = 0; }
System.Action reset = Reset;
n = 5; reset();
__Check((n).ToString(), "0");
