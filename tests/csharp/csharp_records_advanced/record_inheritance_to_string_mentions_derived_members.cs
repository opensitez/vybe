// vybe-test: csharp/csharp_records_advanced/record_inheritance_to_string_mentions_derived_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

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

record Animal(string Name); record Cat(string Name, string Color) : Animal(Name); __P((new Cat("Milo", "Black").ToString().Contains("Color = Black")).ToString());
__Check("True");
