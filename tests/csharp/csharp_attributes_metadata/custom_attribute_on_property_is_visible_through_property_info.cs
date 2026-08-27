// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_property_is_visible_through_property_info
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var property = typeof(Settings).GetProperty("Port");
var attr = (HintAttribute)Attribute.GetCustomAttribute(property, typeof(HintAttribute));
__P((attr.Text).ToString());
__Check("port");

[AttributeUsage(AttributeTargets.Property)] class HintAttribute : Attribute { public string Text { get; } public HintAttribute(string text) { Text = text; } }

class Settings { [Hint("port")] public int Port { get; set; } }

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
