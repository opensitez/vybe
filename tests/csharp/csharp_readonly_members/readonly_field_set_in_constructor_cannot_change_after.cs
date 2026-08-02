// vybe-test: csharp/csharp_readonly_members/readonly_field_set_in_constructor_cannot_change_after
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Immutable{public readonly int Value; public Immutable(int v){Value=v;}}
var obj=new Immutable(42);
__Check((obj.Value).ToString(), "42");
