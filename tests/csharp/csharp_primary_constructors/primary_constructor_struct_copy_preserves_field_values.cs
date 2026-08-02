// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copy_preserves_field_values
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pair(int a, int b) { public int A = a; public int B = b; }
var p = new Pair(2, 3);
var q = p;
__Check((q.A + q.B).ToString(), "5");
