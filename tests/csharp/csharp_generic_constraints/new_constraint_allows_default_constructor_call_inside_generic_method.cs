// vybe-test: csharp/csharp_generic_constraints/new_constraint_allows_default_constructor_call_inside_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Create<T>() where T : new() => new T();
class Widget { public int Value = 42; }
var w = Create<Widget>();
__Check((w.Value).ToString(), "42");
