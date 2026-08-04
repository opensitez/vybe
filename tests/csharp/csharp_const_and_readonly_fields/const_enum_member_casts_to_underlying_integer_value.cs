// vybe-test: csharp/csharp_const_and_readonly_fields/const_enum_member_casts_to_underlying_integer_value
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

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

enum Code { Ok = 0, Err = 1 }
const Code status = Code.Ok;
__P(((int)status).ToString());
__Check("0");
