// vybe-test: csharp/csharp_records_advanced/record_can_define_computed_property_from_primary_members
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

record Rectangle(int Width, int Height) { public int Area => Width * Height; } __P((new Rectangle(3, 7).Area).ToString());
__Check("21");
