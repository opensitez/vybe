// vybe-test: csharp/csharp_init_required_members/required_string_and_optional_init_property_together
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Profile { public required string User; public int Score { get; init; } = 0; }
var p = new Profile { User = "ada", Score = 100 };
__Check((p.User).ToString(), "ada"); __Check((p.Score).ToString(), "100");
