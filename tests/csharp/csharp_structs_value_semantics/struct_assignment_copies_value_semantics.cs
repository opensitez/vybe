// vybe-test: csharp/csharp_structs_value_semantics/struct_assignment_copies_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Counter { public int Value; } var left = new Counter { Value = 1 }; var right = left; right.Value = 9; __Check((left.Value).ToString(), "1"); __Check((right.Value).ToString(), "9");
