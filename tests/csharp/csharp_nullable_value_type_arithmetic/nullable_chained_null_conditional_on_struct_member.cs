// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_chained_null_conditional_on_struct_member
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; public int Y; }
Point? location = new Point { X = 2, Y = 3 };
__Check((location?.X).ToString(), "2");
location = null;
__Check((location?.X ?? -1).ToString(), "-1");
