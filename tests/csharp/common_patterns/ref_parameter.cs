// vybe-test: csharp/common_patterns/ref_parameter
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Ops {
    public static void Double(ref int x) { x *= 2; }
}
int val = 5;
Ops.Double(ref val);
__Check((val).ToString(), "10");
Ops.Double(ref val);
__Check((val).ToString(), "20");
