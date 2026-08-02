// vybe-test: csharp/csharp_scope_variables/multiple_assignment_in_declaration_using_tuple
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a, b) = (3, 7);
__Check((a).ToString(), "3"); __Check((b).ToString(), "7");
