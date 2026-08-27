// vybe-test: csharp/csharp_attributes_metadata/flags_attribute_marks_enum_but_bitwise_result_still_formats_value
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var permission = Permission.Read | Permission.Write;
__P((permission).ToString());
__Check("Read, Write");

[Flags] enum Permission { Read = 1, Write = 2, Execute = 4 }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
