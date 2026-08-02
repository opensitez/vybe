// vybe-test: csharp/csharp_linq_query_syntax/linq_query_selects_names_from_people_with_age_filter
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
class Person {
    public string Name { get; set; }
    public int Age { get; set; }
}
var people = new[] {
    new Person { Name = "Ada", Age = 28 },
    new Person { Name = "Linus", Age = 34 },
    new Person { Name = "Grace", Age = 41 }
};
var names = from p in people
            where p.Age >= 30
            select p.Name;
foreach (var name in names) Console.WriteLine(name);
