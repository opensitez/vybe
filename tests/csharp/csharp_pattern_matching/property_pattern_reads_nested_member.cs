// vybe-test: csharp/csharp_pattern_matching/property_pattern_reads_nested_member
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rect { public int W, H; }
object r = new Rect { W=10, H=5 };
string size = r switch { Rect { W: > 8 } => "wide", _ => "narrow" };
__Check((size).ToString(), "wide");
