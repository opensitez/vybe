// vybe-test: csharp/csharp_tuple_patterns/tuple_deconstruction_assigns_to_multiple_locals
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(string name,int age)=("Alice",30);
__Check((name).ToString(), "Alice"); __Check((age).ToString(), "30");
