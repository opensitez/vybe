// vybe-test: csharp/csharp_json_dom_json_node_tree/node_array_tojsonstring
// Expectations generated from .NET SDK 10 — never hand-written.

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

var a = new System.Text.Json.Nodes.JsonArray();
a.Add(1);
a.Add(2);
__P(a.ToJsonString());
__Check("[1,2]");
