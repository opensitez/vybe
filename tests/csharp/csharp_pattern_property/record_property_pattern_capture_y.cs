// vybe-test: csharp/csharp_pattern_property/record_property_pattern_capture_y
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X,int Y); object o=new Point(3,8); if(o is Point{X:3,Y:var y}) __Check((y).ToString(), "8");
