// vybe-test: csharp/csharp_json_utf8_writer_streaming/utf8_json_writer_case_16

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

var stream = new System.IO.MemoryStream();
using (var writer = new System.Text.Json.Utf8JsonWriter(stream)) {
    writer.WriteStartObject();
    writer.WriteNumber("val", 16);
    writer.WriteEndObject();
}
string str = System.Text.Encoding.UTF8.GetString(stream.ToArray());
__P((str.Contains("16")).ToString());
__Check("True");
