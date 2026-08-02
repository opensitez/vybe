// vybe-test: csharp/csharp_record_types/record_hash_code_equal_for_equal_instances
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Tag(string Name);
var a = new Tag("x"); var b = new Tag("x");
__Check((a.GetHashCode() == b.GetHashCode()).ToString(), "True");
