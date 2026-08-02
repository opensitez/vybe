// vybe-test: csharp/csharp_scope_variables/var_keyword_infers_type_from_right_hand_side
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = "hello";
var number = 42;
__Check((text.GetType().Name).ToString(), "String");
__Check((number.GetType().Name).ToString(), "Int32");
