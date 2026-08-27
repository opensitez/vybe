// vybe-test: csharp/csharp_collections_name_value_collection/name_value_collection_case_1

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

var nvc = new System.Collections.Specialized.NameValueCollection();
nvc.Add("header_1", "val1");
nvc.Add("header_1", "val2");
__P(nvc["header_1"]);
__Check("val1,val2");
