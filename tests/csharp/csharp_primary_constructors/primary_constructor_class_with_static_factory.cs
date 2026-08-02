// vybe-test: csharp/csharp_primary_constructors/primary_constructor_class_with_static_factory
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token(int id) {
    public int Id => id;
    public static Token Default() => new Token(0);
}
__Check((Token.Default().Id).ToString(), "0");
