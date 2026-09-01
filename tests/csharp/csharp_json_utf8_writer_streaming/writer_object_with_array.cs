// vybe-test: csharp/csharp_json_utf8_writer_streaming/writer_object_with_array
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

var stream = new System.IO.MemoryStream();
using (var w = new System.Text.Json.Utf8JsonWriter(stream)) {
  w.WriteStartObject();
  w.WriteStartArray("xs");
  w.WriteNumberValue(7);
  w.WriteEndArray();
  w.WriteEndObject();
}
__P(System.Text.Encoding.UTF8.GetString(stream.ToArray()));
__Check("{\"xs\":[7]}");
