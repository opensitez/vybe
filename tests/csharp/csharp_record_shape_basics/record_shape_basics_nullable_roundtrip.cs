// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
int? maybe = 39; __Check((maybe.HasValue && maybe.Value == 39).ToString(), "True");
