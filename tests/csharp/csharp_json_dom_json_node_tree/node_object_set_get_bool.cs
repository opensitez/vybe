// vybe-test: csharp/csharp_json_dom_json_node_tree/node_object_set_get_bool
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

var o = new System.Text.Json.Nodes.JsonObject();
o["b"] = true;
__P(o["b"].ToString());
__Check("true");
