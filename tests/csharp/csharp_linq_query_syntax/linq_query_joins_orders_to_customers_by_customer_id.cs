// vybe-test: csharp/csharp_linq_query_syntax/linq_query_joins_orders_to_customers_by_customer_id
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

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
foreach (var item in query) __P((item).ToString());
__Check("Ada:Pen\nGrace:Book");
