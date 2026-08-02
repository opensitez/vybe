// vybe-test: csharp/csharp_oop_inheritance/derived_class_can_have_additional_members
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Vehicle { public int Wheels = 4; }
class Bike : Vehicle { public bool HasKickstand = true; }
var bike = new Bike();
__Check((bike.Wheels).ToString(), "4");
__Check((bike.HasKickstand).ToString(), "True");
