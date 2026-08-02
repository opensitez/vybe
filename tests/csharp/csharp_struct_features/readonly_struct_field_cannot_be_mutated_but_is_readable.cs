// vybe-test: csharp/csharp_struct_features/readonly_struct_field_cannot_be_mutated_but_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Immutable { public readonly int Value; public Immutable(int v) { Value=v; } }
var obj = new Immutable(7);
__Check((obj.Value).ToString(), "7");
