// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_struct_field_mutation
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Canvas{public struct Dot{public int X;} public Dot Make(){var d=new Dot(); d.X=9; return d;}} __Check((new Canvas().Make().X).ToString(), "9");
