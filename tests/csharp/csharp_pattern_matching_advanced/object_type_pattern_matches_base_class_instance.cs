// vybe-test: csharp/csharp_pattern_matching_advanced/object_type_pattern_matches_base_class_instance
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { } class Dog : Animal { } object pet = new Dog(); __Check((pet is Animal).ToString(), "True");
