// vybe-test: csharp/csharp_record_types/record_struct_is_a_value_type
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Coord(double Lat, double Lon);
var a = new Coord(1.0, 2.0);
var b = a;
__Check((a == b).ToString(), "True");
