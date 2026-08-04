// vybe-test: csharp/csharp_string_builder/remove_deletes_character_range_by_start_and_count
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

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

var sb = new System.Text.StringBuilder("hello");
sb.Remove(1,3);
__P((sb.ToString()).ToString());
__Check("ho");
