// vybe-test: csharp/csharp_init_required_members/init_property_guid_value_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Ref { public System.Guid Id { get; init; } }
var id = new System.Guid("11111111-1111-1111-1111-111111111111");
var r = new Ref { Id = id };
__Check((r.Id == id).ToString(), "True");
