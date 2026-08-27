// vybe-test: csharp/csharp_oop_record_class_sealed_tostring/record_sealed_tostring_case_1

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

var item = new RecordItem_1("ID_1");
__P(item.ToString());
__Check("Custom_ID_1");

record RecordItem_1(string Tag) {
    public sealed override string ToString() => $"Custom_{Tag}";
}
