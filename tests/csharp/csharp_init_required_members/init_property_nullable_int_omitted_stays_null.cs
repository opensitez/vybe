// vybe-test: csharp/csharp_init_required_members/init_property_nullable_int_omitted_stays_null
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Maybe { public int? Count { get; init; } }
var m = new Maybe();
__Check((m.Count.HasValue).ToString(), "False");
