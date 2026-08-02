// vybe-test: csharp/csharp_init_required_members/init_property_array_reference_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bundle { public int[] Items { get; init; } = new int[0]; }
var b = new Bundle { Items = new[] { 1, 2, 3 } };
__Check((b.Items.Length).ToString(), "3");
