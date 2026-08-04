// vybe-test: csharp/csharp_structs_value_semantics/nullable_struct_value_exposes_underlying_member
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

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

System.DateTime? value = new System.DateTime(2024, 1, 1); __P((value.Value.Year).ToString());
__Check("2024");
