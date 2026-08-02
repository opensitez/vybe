// vybe-test: csharp/csharp_readonly_members/const_local_not_changeable_but_usable_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

const int MAX=100;
__Check((MAX*2).ToString(), "200");
