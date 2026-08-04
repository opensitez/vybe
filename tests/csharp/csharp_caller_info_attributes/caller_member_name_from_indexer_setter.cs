// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_indexer_setter
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

class Row {
    int[] cells = new int[3];
    public int this[int i] {
        set {
            LogWrite();
            cells[i] = value;
        }
    }
    void LogWrite([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
var r = new Row(); r[0] = 7; __P((r[0]).ToString());
__Check("Item\n7");
