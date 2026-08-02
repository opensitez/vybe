// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_constructor
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node {
    public Node() { Trace(); }
    void Trace([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), ".ctor");
}
new Node();
