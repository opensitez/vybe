// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_char_int_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.Dictionary<char, int> map = new();
map['A'] = 1;
__Check((map['A']).ToString(), "1");
