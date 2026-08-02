// vybe-test: csharp/csharp_generic_constraints/interface_constraint_enforces_method_availability_at_compile_time
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILabel { string Label(); }
class Tag : ILabel { public string Label() => "tag"; }
string Get<T>(T t) where T : ILabel => t.Label();
__Check((Get(new Tag())).ToString(), "tag");
