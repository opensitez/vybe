// vybe-test: csharp/csharp_struct_copy_independence/struct_field_mutation_on_copy_leaves_original_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_struct_copy_independence.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; }
var left = new Point { X = 1 };
var right = left;
right.X = 9;
__Check((left.X).ToString(), "1");
__Check((right.X).ToString(), "9");
