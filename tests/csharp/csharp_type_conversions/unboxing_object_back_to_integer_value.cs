// vybe-test: csharp/csharp_type_conversions/unboxing_object_back_to_integer_value
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object boxed = 21; int count = (int)boxed; __Check((count + 1).ToString(), "22");
