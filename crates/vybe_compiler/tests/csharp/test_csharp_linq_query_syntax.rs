use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    linq_query_selects_names_from_people_with_age_filter,
    r#"
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
"#,
    ["Linus", "Grace"]
);

csharp_case!(
    linq_query_filters_even_numbers_then_projects_squares,
    r#"
using System.Linq;
var values = new[] { 1, 2, 3, 4, 5, 6 };
var query = from value in values
            where value % 2 == 0
            select value * value;
foreach (var value in query) Console.WriteLine(value);
"#,
    ["4", "16", "36"]
);

csharp_case!(
    linq_query_orders_words_by_length_then_name,
    r#"
using System.Linq;
var words = new[] { "pear", "fig", "banana", "kiwi" };
var query = from word in words
            orderby word.Length, word
            select word;
foreach (var word in query) Console.WriteLine(word);
"#,
    ["fig", "kiwi", "pear", "banana"]
);

csharp_case!(
    linq_query_uses_let_clause_for_trimmed_values,
    r#"
using System.Linq;
var raw = new[] { "  alpha  ", " beta", "gamma " };
var query = from value in raw
            let trimmed = value.Trim()
            select trimmed + ":" + trimmed.Length;
foreach (var item in query) Console.WriteLine(item);
"#,
    ["alpha:5", "beta:4", "gamma:5"]
);

csharp_case!(
    linq_query_uses_multiple_from_clauses_to_flatten_pairs,
    r#"
using System.Linq;
var prefixes = new[] { "A", "B" };
var suffixes = new[] { 1, 2, 3 };
var query = from prefix in prefixes
            from suffix in suffixes
            where suffix != 2
            select prefix + suffix;
foreach (var item in query) Console.WriteLine(item);
"#,
    ["A1", "A3", "B1", "B3"]
);

csharp_case!(
    linq_query_groups_words_by_first_letter,
    r#"
using System.Linq;
var words = new[] { "apple", "ant", "banana", "boat" };
var groups = from word in words
             group word by word[0] into grouped
             orderby grouped.Key
             select grouped;
foreach (var group in groups) {
    Console.WriteLine(group.Key);
    Console.WriteLine(group.Count());
}
"#,
    ["a", "2", "b", "2"]
);

csharp_case!(
    linq_query_joins_orders_to_customers_by_customer_id,
    r#"
using System.Linq;
class Customer {
    public int Id { get; set; }
    public string Name { get; set; }
}
class Order {
    public int CustomerId { get; set; }
    public string Item { get; set; }
}
var customers = new[] {
    new Customer { Id = 1, Name = "Ada" },
    new Customer { Id = 2, Name = "Grace" }
};
var orders = new[] {
    new Order { CustomerId = 2, Item = "Book" },
    new Order { CustomerId = 1, Item = "Pen" }
};
var query = from customer in customers
            join order in orders on customer.Id equals order.CustomerId
            orderby customer.Name
            select customer.Name + ":" + order.Item;
foreach (var item in query) Console.WriteLine(item);
"#,
    ["Ada:Pen", "Grace:Book"]
);

csharp_case!(
    linq_query_uses_group_into_then_filters_large_groups,
    r#"
using System.Linq;
var words = new[] { "ape", "ant", "boat", "berry", "cat" };
var query = from word in words
            group word by word.Length into groups
            where groups.Count() >= 2
            orderby groups.Key
            select groups.Key + ":" + groups.Count();
foreach (var item in query) Console.WriteLine(item);
"#,
    ["3:3"]
);

csharp_case!(
    linq_query_projects_anonymous_type_fields,
    r#"
using System.Linq;
class TaggedTotal {
    public string Label { get; set; }
    public int Total { get; set; }
}
var prices = new[] { 3, 7, 10 };
var query = from price in prices
            select new TaggedTotal { Label = "item", Total = price * 2 };
foreach (var item in query) Console.WriteLine(item.Label + ":" + item.Total);
"#,
    ["item:6", "item:14", "item:20"]
);

csharp_case!(
    linq_query_combines_where_orderby_and_select,
    r#"
using System.Linq;
var values = new[] { 9, 1, 6, 2, 3 };
var query = from value in values
            where value >= 3
            orderby value descending
            select value - 1;
foreach (var item in query) Console.WriteLine(item);
"#,
    ["8", "5", "2"]
);
