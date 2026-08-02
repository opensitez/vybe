// vybe-test: csharp/winforms/new_point_properties
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = new Point(100, 200);
        __Check((p.x).ToString(), "100");
        __Check((p.y).ToString(), "200");
