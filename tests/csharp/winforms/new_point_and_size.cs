// vybe-test: csharp/winforms/new_point_and_size
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = new Point(10, 20);
        var s = new Size(100, 50);
        __Check((p.x + " " + p.y).ToString(), "10 20");
        __Check((s.width + " " + s.height).ToString(), "100 50");
