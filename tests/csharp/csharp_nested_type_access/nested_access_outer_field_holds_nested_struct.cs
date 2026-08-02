// vybe-test: csharp/csharp_nested_type_access/nested_access_outer_field_holds_nested_struct
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Grid{public struct Cell{public int V;} Cell _c; public Grid(){_c.V=6;} public int Read()=>_c.V;} __Check((new Grid().Read()).ToString(), "6");
