// vybe-test: csharp/csharp_properties_accessors/auto_property_initializer_sets_default_title
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Article {
    public string Title { get; set; } = "draft";
}
var article = new Article();
__P((article.Title).ToString());
article.Title = "published";
__P((article.Title).ToString());
__Check("draft\npublished");
