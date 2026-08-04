// vybe-test: csharp/csharp_enum_operations/enum_parse_converts_string_name_to_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

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

enum Day{Mon,Tue,Wed,Thu,Fri}
var d = (Day)System.Enum.Parse(typeof(Day),"Wed");
__P((d).ToString());
__Check("Wed");
