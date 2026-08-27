// vybe-test: csharp/csharp_json_utf8_reader_streaming/utf8_json_reader_case_17

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

byte[] json = System.Text.Encoding.UTF8.GetBytes("{\"id\":17}");
var reader = new System.Text.Json.Utf8JsonReader(json);
bool ok = reader.Read();
__P(ok.ToString());
__P(reader.TokenType.ToString());
__Check("True\nStartObject");
