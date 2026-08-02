// vybe-test: csharp/csharp_static_type_behaviors/static_property_tracks_number_of_created_objects
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token {
    public static int Created { get; private set; }
    public Token() { Created++; }
}
new Token();
new Token();
new Token();
__Check((Token.Created).ToString(), "3");
