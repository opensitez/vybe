// vybe-test: csharp/csharp_primary_constructors/primary_constructor_two_instances_are_independent
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Slot(int id) { public int Id => id; }
var a = new Slot(1);
var b = new Slot(2);
__Check((a.Id).ToString(), "1"); __Check((b.Id).ToString(), "2");
