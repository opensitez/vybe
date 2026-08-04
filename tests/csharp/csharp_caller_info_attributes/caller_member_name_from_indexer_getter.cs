// vybe-test: csharp/csharp_caller_info_attributes/caller_member_name_from_indexer_getter
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
    int[] cells = { 1, 2, 3 };
    public int this[int i] {
        get {
            LogAccess();
            return cells[i];
        }
    }
    void LogAccess([System.Runtime.CompilerServices.CallerMemberName] string member = "") => __P((member).ToString());
}
__P((new Row()[1]).ToString());
__Check("Item\n2");
