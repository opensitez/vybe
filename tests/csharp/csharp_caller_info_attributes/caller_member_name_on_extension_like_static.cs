// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_extension_like_static
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Ext {
    public static void Dump(this string s, [System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "<Main>$");
}
"hi".Dump();
