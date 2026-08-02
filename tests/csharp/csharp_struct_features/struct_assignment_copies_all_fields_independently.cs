// vybe-test: csharp/csharp_struct_features/struct_assignment_copies_all_fields_independently
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X, Y; }
var a = new Point { X=1, Y=2 };
var b = a;
b.X = 99;
__Check((a.X).ToString(), "1");
