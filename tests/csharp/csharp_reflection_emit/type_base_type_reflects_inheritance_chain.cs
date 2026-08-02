// vybe-test: csharp/csharp_reflection_emit/type_base_type_reflects_inheritance_chain
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{} class B:A{} class C:B{}
__Check((typeof(C).BaseType.Name).ToString(), "B");
__Check((typeof(C).BaseType.BaseType.Name).ToString(), "A");
