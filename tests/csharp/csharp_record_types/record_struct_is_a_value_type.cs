// vybe-test: csharp/csharp_record_types/record_struct_is_a_value_type
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

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

record struct Coord(double Lat, double Lon);
var a = new Coord(1.0, 2.0);
var b = a;
__P((a == b).ToString());
__Check("True");
