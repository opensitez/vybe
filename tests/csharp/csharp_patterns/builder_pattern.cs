// vybe-test: csharp/csharp_patterns/builder_pattern
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class QueryBuilder {
    private string query = "SELECT *";
    public QueryBuilder From(string table) {
        query += " FROM " + table;
        return this;
    }
    public QueryBuilder Where(string condition) {
        query += " WHERE " + condition;
        return this;
    }
    public string Build() { return query; }
}
var q = new QueryBuilder().From("users").Where("age > 18").Build();
__Check((q).ToString(), "SELECT * FROM users WHERE age > 18");
