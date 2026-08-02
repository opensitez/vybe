// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_predicate_from_method_group
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static bool IsEven(int n) => n % 2 == 0;
System.Predicate<int> even = IsEven;
__Check((even(4)).ToString(), "True"); __Check((even(3)).ToString(), "False");
