// vybe-test: csharp/csharp_object_equality/string_equals_compares_content_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
__Check((a.Equals(b)).ToString(), "True");
