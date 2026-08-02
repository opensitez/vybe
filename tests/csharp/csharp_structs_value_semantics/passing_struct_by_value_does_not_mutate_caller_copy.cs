// vybe-test: csharp/csharp_structs_value_semantics/passing_struct_by_value_does_not_mutate_caller_copy
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Counter { public int Value; } void Bump(Counter counter) { counter.Value++; } var counter = new Counter { Value = 2 }; Bump(counter); __Check((counter.Value).ToString(), "2");
