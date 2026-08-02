// vybe-test: csharp/csharp_init_required_members/init_property_list_reference_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder { public System.Collections.Generic.List<int> Values { get; init; } = new(); }
var h = new Holder { Values = new System.Collections.Generic.List<int> { 4, 5 } };
__Check((h.Values.Count).ToString(), "2");
