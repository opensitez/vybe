// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_extension_like_static
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

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

static class Ext {
    public static void Dump(this string s, [System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
"hi".Dump();
__Check("<Main>$");
