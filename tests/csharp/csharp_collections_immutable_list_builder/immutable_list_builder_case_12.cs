// vybe-test: csharp/csharp_collections_immutable_list_builder/immutable_list_builder_case_12

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

var b = System.Collections.Immutable.ImmutableList.CreateBuilder<string>();
b.Add("Item_12");
var list = b.ToImmutable();
__P(list.Count.ToString());
__P(list[0]);
__Check("1\nItem_12");
