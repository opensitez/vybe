// vybe-test: csharp/csharp_reflection_emit/type_is_assignable_from_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{} class B:A{}
__Check((typeof(A).IsAssignableFrom(typeof(B))).ToString(), "True");
__Check((typeof(B).IsAssignableFrom(typeof(A))).ToString(), "False");
