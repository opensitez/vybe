// vybe-test: csharp/csharp_pattern_property/inheritance_property_pattern_on_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { public string Kind; } class Dog : Animal { public int Legs; } object o=new Dog{Kind="pet",Legs=4}; __Check((o is Dog{Legs:4,Kind:"pet"}).ToString(), "True");
