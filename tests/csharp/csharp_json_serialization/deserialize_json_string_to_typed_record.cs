// vybe-test: csharp/csharp_json_serialization/deserialize_json_string_to_typed_record
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

using static __Harness;

var opts = new System.Text.Json.JsonSerializerOptions { TypeInfoResolver = new System.Text.Json.Serialization.Metadata.DefaultJsonTypeInfoResolver() };
var doc = System.Text.Json.JsonDocument.Parse("{\"id\": 1}");
__P((doc.RootElement.GetProperty("id").GetInt32() == 1).ToString());
__Check("True");
public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
