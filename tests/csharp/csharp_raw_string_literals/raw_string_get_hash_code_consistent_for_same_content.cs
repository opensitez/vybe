// vybe-test: csharp/csharp_raw_string_literals/raw_string_get_hash_code_consistent_for_same_content
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a="""hash"""; string b="""hash"""; __Check((a.GetHashCode()==b.GetHashCode()).ToString(), "True");
