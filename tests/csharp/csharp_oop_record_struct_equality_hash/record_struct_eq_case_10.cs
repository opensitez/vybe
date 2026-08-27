// vybe-test: csharp/csharp_oop_record_struct_equality_hash/record_struct_eq_case_10

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

var p1 = new ItemPoint_10(10, 20);
var p2 = new ItemPoint_10(10, 20);
__P((p1 == p2).ToString());
__P((p1.GetHashCode() == p2.GetHashCode()).ToString());
__Check("True\nTrue");

record struct ItemPoint_10(int X, int Y);
