// vybe-test: csharp/csharp_json_custom_converter_polymorphic/options_serialize_array
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

__P(System.Text.Json.JsonSerializer.Serialize(new int[] { 1, 2, 3 }));
__Check("[1,2,3]");
