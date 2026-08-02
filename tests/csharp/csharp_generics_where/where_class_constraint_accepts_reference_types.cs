// vybe-test: csharp/csharp_generics_where/where_class_constraint_accepts_reference_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Wrap<T>(T v) where T:class=>v;
__Check((Wrap("hello")).ToString(), "hello");
