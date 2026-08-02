// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_char_to_int_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<char, int> code = c => (int)c;
__Check((code('A')).ToString(), "65");
