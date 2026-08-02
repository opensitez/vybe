// vybe-test: csharp/csharp_primary_constructors/primary_constructor_method_mutates_field_not_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Acc(int start) { int total = start; public void Add(int n) { total += n; } public int Value => total; }
var a = new Acc(1);
a.Add(4);
__Check((a.Value).ToString(), "5");
