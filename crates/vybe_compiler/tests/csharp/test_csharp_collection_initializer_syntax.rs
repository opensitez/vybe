//! Collection and object initializer syntax at construction sites.
use super::helpers::run_csharp;

#[test]
fn list_initializer_populates_elements_in_source_order() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var items = new List<int> { 3, 1, 4 };
Console.WriteLine(items[0]);
Console.WriteLine(items[2]);
"#
        ),
        &["3", "4"]
    );
}

#[test]
fn dictionary_initializer_binds_keys_to_values() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var map = new Dictionary<string, int> { ["x"] = 9, ["y"] = 2 };
Console.WriteLine(map["y"]);
"#
        ),
        &["2"]
    );
}

#[test]
fn array_initializer_literal_sets_length_and_elements() {
    assert_eq!(
        run_csharp(
            r#"
var data = new[] { 10, 20, 30 };
Console.WriteLine(data.Length);
Console.WriteLine(data[1]);
"#
        ),
        &["3", "20"]
    );
}

#[test]
fn object_initializer_sets_public_fields_before_use() {
    assert_eq!(
        run_csharp(
            r#"
class Point { public int X; public int Y; }
var point = new Point { X = 2, Y = 5 };
Console.WriteLine(point.X + point.Y);
"#
        ),
        &["7"]
    );
}

#[test]
fn object_initializer_can_invoke_properties() {
    assert_eq!(
        run_csharp(
            r#"
class User { public string Name { get; set; } }
var user = new User { Name = "Ada" };
Console.WriteLine(user.Name);
"#
        ),
        &["Ada"]
    );
}

#[test]
fn nested_object_initializer_builds_graph_in_one_expression() {
    assert_eq!(
        run_csharp(
            r#"
class Address { public string City { get; set; } }
class Person { public Address Home { get; set; } }
var person = new Person { Home = new Address { City = "Oslo" } };
Console.WriteLine(person.Home.City);
"#
        ),
        &["Oslo"]
    );
}

#[test]
fn hashset_initializer_collection_adds_unique_members() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var set = new HashSet<int> { 2, 3, 2 };
Console.WriteLine(set.Count);
"#
        ),
        &["2"]
    );
}

#[test]
fn jagged_array_initializer_creates_rows_with_varying_lengths() {
    assert_eq!(
        run_csharp(
            r#"
int[][] grid = {
    new[] { 1, 2 },
    new[] { 3, 4, 5 }
};
Console.WriteLine(grid[1].Length);
Console.WriteLine(grid[1][2]);
"#
        ),
        &["3", "5"]
    );
}

#[test]
fn list_initializer_after_empty_constructor_appends_in_order() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var items = new List<string>();
items.Add("first");
items.Add("second");
Console.WriteLine(items[1]);
"#
        ),
        &["second"]
    );
}

#[test]
fn readonly_struct_initializer_sets_init_only_properties() {
    assert_eq!(
        run_csharp(
            r#"
readonly struct Token {
    public string Value { get; init; }
}
var token = new Token { Value = "abc" };
Console.WriteLine(token.Value);
"#
        ),
        &["abc"]
    );
}

#[test]
fn queue_initializer_enqueues_elements_for_fifo_order() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
Console.WriteLine(queue.Dequeue());
"#
        ),
        &["1"]
    );
}

#[test]
fn stack_initializer_pushes_elements_for_lifo_order() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var stack = new Stack<int>();
stack.Push(1);
stack.Push(2);
Console.WriteLine(stack.Pop());
"#
        ),
        &["2"]
    );
}
