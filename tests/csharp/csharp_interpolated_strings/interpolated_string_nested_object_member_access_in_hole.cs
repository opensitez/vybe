// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_nested_object_member_access_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public int A; public int B; }
var pair = new Pair { A = 2, B = 3 };
__Check(($"{pair.A + pair.B}").ToString(), "5");
