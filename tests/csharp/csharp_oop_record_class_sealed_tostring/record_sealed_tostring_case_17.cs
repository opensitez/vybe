// vybe-test: csharp/csharp_oop_record_class_sealed_tostring/record_sealed_tostring_case_17

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var item = new RecordItem_17("ID_17");
__P(item.ToString());
__Check("Custom_ID_17");

record RecordItem_17(string Tag) {
    public sealed override string ToString() => $"Custom_{Tag}";
}
