// vybe-test: csharp/csharp_conditional_expressions/null_coalescing_assignment_sets_only_when_null
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a=null; a??="assigned";
string b="existing"; b??="assigned";
__Check((a).ToString(), "assigned"); __Check((b).ToString(), "existing");
