// vybe-test: csharp/csharp_pattern_matching_advanced/property_pattern_like_access_via_guarded_type_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X { get; set; } public int Y { get; set; } } object item = new Point { X = 5, Y = 8 }; if (item is Point point && point.X == 5) __Check((point.Y).ToString(), "8");
