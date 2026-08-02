// vybe-test: csharp/csharp_oop/enum_usage
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Direction { North, South, East, West }
Direction d = Direction.East;
__Check((d).ToString(), "East");
