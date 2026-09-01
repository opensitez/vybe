// vybe-test: csharp/csharp_json_utf8_reader_streaming/reader_sum_numbers
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

byte[] bytes = System.Text.Encoding.UTF8.GetBytes("{\"id\":1,\"nm\":\"a\",\"ok\":true,\"xs\":[1,2]}");
var r = new System.Text.Json.Utf8JsonReader(bytes);
int s = 0;
while (r.Read()) if (r.TokenType == System.Text.Json.JsonTokenType.Number) s += r.GetInt32();
__P(s.ToString());
__Check("4");
