// vybe-test: csharp/csharp_readonly_members/const_field_accessible_without_instance_on_type
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Physics{public const double C=299792458.0;}
__Check((Physics.C>0).ToString(), "True");
