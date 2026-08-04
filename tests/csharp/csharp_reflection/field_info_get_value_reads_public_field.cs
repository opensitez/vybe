// vybe-test: csharp/csharp_reflection/field_info_get_value_reads_public_field
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

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

class Data { public int X = 3; }
var obj = new Data();
var field = typeof(Data).GetField("X");
__P((field.GetValue(obj)).ToString());
__Check("3");
