// vybe-test: csharp/csharp_reflection_constructor_info_invoke/constructor_info_case_6

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

var ctor = typeof(System.Text.StringBuilder).GetConstructor(new Type[] { typeof(string) });
var sb = (System.Text.StringBuilder)ctor.Invoke(new object[] { "Item_6" });
__P(sb.ToString());
__Check("Item_6");
