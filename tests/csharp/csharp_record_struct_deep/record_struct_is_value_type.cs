// vybe-test: csharp/csharp_record_struct_deep/record_struct_is_value_type
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Coord(int X,int Y); __Check((typeof(Coord).IsValueType).ToString(), "True");
