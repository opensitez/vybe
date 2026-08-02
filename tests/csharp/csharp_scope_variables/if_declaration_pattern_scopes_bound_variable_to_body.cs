// vybe-test: csharp/csharp_scope_variables/if_declaration_pattern_scopes_bound_variable_to_body
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = "scoped";
if(o is string text)
    __Check((text.Length).ToString(), "6");
__Check(("done").ToString(), "done");
