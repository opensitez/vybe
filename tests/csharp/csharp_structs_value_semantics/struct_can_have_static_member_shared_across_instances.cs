// vybe-test: csharp/csharp_structs_value_semantics/struct_can_have_static_member_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Token { public static int Count; public Token(int _) { Count++; } } new Token(1); new Token(2); __Check((Token.Count).ToString(), "2");
