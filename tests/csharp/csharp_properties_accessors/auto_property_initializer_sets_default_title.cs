// vybe-test: csharp/csharp_properties_accessors/auto_property_initializer_sets_default_title
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var article = new Article();
__P((article.Title).ToString());
article.Title = "published";
__P((article.Title).ToString());
__Check("draft\npublished");

class Article {
    public string Title { get; set; } = "draft";
}

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
