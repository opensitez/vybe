// vybe-test: csharp/csharp_properties_accessors/auto_property_initializer_sets_default_title
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Article {
    public string Title { get; set; } = "draft";
}
var article = new Article();
__Check((article.Title).ToString(), "draft");
article.Title = "published";
__Check((article.Title).ToString(), "published");
