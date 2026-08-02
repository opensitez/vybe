// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_static_method_group_to_action
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int hits = 0;
static void Hit() { hits++; }
System.Action strike = Hit;
strike(); strike();
__Check((hits).ToString(), "2");
