// vybe-test: csharp/csharp_linq_query_syntax/linq_query_selects_names_from_people_with_age_filter
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using static __Harness;
using System.Linq;

var people = new[] {
    new Person { Name = "Ada", Age = 28 },
    new Person { Name = "Linus", Age = 34 },
    new Person { Name = "Grace", Age = 41 }
}
;
var names = from p in people
            where p.Age >= 30
            select p.Name;
foreach (var name in names) __P((name).ToString());
__Check("Linus\nGrace");

class Person {
    public string Name { get; set; }
    public int Age { get; set; }
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
