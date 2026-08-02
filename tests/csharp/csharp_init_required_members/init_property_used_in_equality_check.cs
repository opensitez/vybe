// vybe-test: csharp/csharp_init_required_members/init_property_used_in_equality_check
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tag { public string Name { get; init; } = ""; }
var a = new Tag { Name = "x" };
var b = new Tag { Name = "x" };
__Check((a.Name == b.Name).ToString(), "True");
