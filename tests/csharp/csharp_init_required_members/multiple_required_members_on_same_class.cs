// vybe-test: csharp/csharp_init_required_members/multiple_required_members_on_same_class
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public required int Left; public required int Right; }
var p = new Pair { Left = 3, Right = 9 };
__Check((p.Left + p.Right).ToString(), "12");
