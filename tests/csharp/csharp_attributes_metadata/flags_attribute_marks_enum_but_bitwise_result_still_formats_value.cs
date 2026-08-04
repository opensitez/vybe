// vybe-test: csharp/csharp_attributes_metadata/flags_attribute_marks_enum_but_bitwise_result_still_formats_value
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

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

using System; [Flags] enum Permission { Read = 1, Write = 2, Execute = 4 } var permission = Permission.Read | Permission.Write; __P((permission).ToString());
__Check("Read, Write");
