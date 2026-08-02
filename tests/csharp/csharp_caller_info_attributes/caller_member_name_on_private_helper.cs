// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_on_private_helper
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Service {
    public int Compute() => Helper();
    int Helper([System.Runtime.CompilerServices.CallerMemberName] string member = "") {
        __Check((member).ToString(), "Compute");
        return 1;
    }
}
__Check((new Service().Compute()).ToString(), "1");
