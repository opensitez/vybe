// vybe-test: csharp/csharp_linq_groupby_join/order_by_then_by_applies_secondary_sort_on_equal_primary_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

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

var items = new[] { (Name:"b",Age:2),(Name:"a",Age:3),(Name:"a",Age:1) };
var sorted = items.OrderBy(x => x.Name).ThenBy(x => x.Age);
foreach (var x in sorted) __P(($"{x.Name}{x.Age}").ToString());
__Check("a1\na3\nb2");
