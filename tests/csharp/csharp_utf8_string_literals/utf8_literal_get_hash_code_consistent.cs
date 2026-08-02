// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_get_hash_code_consistent
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=u8"hash"; var b=u8"hash"; __Check((a.GetHashCode()==b.GetHashCode()).ToString(), "True");
