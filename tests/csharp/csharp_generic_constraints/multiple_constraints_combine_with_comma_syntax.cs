// vybe-test: csharp/csharp_generic_constraints/multiple_constraints_combine_with_comma_syntax
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IName { string Name(); }
T Make<T>() where T : IName, new() => new T();
class Item : IName { public string Name() => "item"; }
__Check((Make<Item>().Name()).ToString(), "item");
