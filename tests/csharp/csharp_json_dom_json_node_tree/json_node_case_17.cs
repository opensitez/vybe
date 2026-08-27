// vybe-test: csharp/csharp_json_dom_json_node_tree/json_node_case_17

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

var obj = new System.Text.Json.Nodes.JsonObject();
obj["num"] = 17;
__P(obj["num"]?.ToString());
__Check("17");
