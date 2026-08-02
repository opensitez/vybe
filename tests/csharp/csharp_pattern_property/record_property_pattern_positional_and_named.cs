// vybe-test: csharp/csharp_pattern_property/record_property_pattern_positional_and_named
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X,int Y); object o=new Point(1,2); __Check((o is Point{X:1,Y:2}).ToString(), "True");
