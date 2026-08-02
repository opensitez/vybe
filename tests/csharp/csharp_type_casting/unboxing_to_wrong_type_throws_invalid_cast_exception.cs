// vybe-test: csharp/csharp_type_casting/unboxing_to_wrong_type_throws_invalid_cast_exception
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object boxed = 42;
string result = "";
try { string s = (string)boxed; }
catch(System.InvalidCastException) { result = "bad"; }
__Check((result).ToString(), "bad");
