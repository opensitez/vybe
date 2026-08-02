// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    public T Read<T>([System.Runtime.CompilerServices.CallerMemberName] string member = "") {
        __Check((member).ToString(), "Read");
        return default(T);
    }
}
new Box().Read<int>();
