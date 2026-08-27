// vybe-test: csharp/csharp_patterns/builder_pattern
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var q = new QueryBuilder().From("users").Where("age > 18").Build();
__P((q).ToString());
__Check("SELECT * FROM users WHERE age > 18");

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
