// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_indexer_setter
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Row {
    int[] cells = new int[3];
    public int this[int i] {
        set {
            LogWrite();
            cells[i] = value;
        }
    }
    void LogWrite([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __Check((member).ToString(), "Item");
}
var r = new Row(); r[0] = 7; __Check((r[0]).ToString(), "7");
