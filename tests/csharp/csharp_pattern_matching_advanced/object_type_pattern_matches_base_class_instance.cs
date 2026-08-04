// vybe-test: csharp/csharp_pattern_matching_advanced/object_type_pattern_matches_base_class_instance
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Animal { } class Dog : Animal { } object pet = new Dog(); __P((pet is Animal).ToString());
__Check("True");
