// vybe-test: csharp/csharp_enum_operations/enum_tostring_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Status{Pending,Active,Done} __Check((Status.Active.ToString()).ToString(), "Active");
