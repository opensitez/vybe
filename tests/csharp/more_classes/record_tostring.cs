// vybe-test: csharp/more_classes/record_tostring
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y) {
            public string Display() {
                return "Point(" + X + ", " + Y + ")";
            }
        }
        var p = new Point(3, 7);
        __Check((p.Display()).ToString(), "Point(3, 7)");
