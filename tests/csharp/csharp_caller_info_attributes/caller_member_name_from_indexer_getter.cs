// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_indexer_getter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Row {
    int[] cells = { 1, 2, 3 };
    public int this[int i] {
        get {
            LogAccess();
            return cells[i];
        }
    }
    void LogAccess([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Item");
}
__Check((new Row()[1]).ToString(), "2");
