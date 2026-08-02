// vybe-test: csharp/csharp_nested_type_member_access/nested_static_class_reads_outer_static_private_state
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    static int tally = 3;
    static class Inner {
        public static int Read() { return tally; }
    }
    public static int Via() { return Inner.Read(); }
}
__Check((Outer.Via()).ToString(), "3");
