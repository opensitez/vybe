//! Target-typed `new()`, target-typed delegates, and `Func`/`Action` inference.

csharp_cases! {
    target_new_list_int_inferred_from_variable_type => {
        r#"System.Collections.Generic.List<int> values = new();
values.Add(7);
Console.WriteLine(values[0]);"#,
        ["7"]
    };

    target_new_list_string_inferred_from_variable_type => {
        r#"System.Collections.Generic.List<string> names = new();
names.Add("Ada");
Console.WriteLine(names[0]);"#,
        ["Ada"]
    };

    target_new_dictionary_inferred_key_value_types => {
        r#"System.Collections.Generic.Dictionary<string, int> map = new();
map["count"] = 3;
Console.WriteLine(map["count"]);"#,
        ["3"]
    };

    target_new_hashset_inferred_element_type => {
        r#"System.Collections.Generic.HashSet<int> set = new();
set.Add(1); set.Add(1); set.Add(2);
Console.WriteLine(set.Count);"#,
        ["2"]
    };

    target_new_queue_inferred_element_type => {
        r#"System.Collections.Generic.Queue<int> q = new();
q.Enqueue(4); q.Enqueue(5);
Console.WriteLine(q.Dequeue()); Console.WriteLine(q.Dequeue());"#,
        ["4", "5"]
    };

    target_new_stack_inferred_element_type => {
        r#"System.Collections.Generic.Stack<int> s = new();
s.Push(1); s.Push(2);
Console.WriteLine(s.Pop()); Console.WriteLine(s.Pop());"#,
        ["2", "1"]
    };

    target_new_linked_list_inferred_element_type => {
        r#"System.Collections.Generic.LinkedList<int> list = new();
list.AddLast(10); list.AddLast(20);
Console.WriteLine(list.First.Value); Console.WriteLine(list.Last.Value);"#,
        ["10", "20"]
    };

    target_new_list_bool_inferred => {
        r#"System.Collections.Generic.List<bool> flags = new();
flags.Add(true); flags.Add(false);
Console.WriteLine(flags[0]); Console.WriteLine(flags[1]);"#,
        ["True", "False"]
    };

    target_new_list_double_inferred => {
        r#"System.Collections.Generic.List<double> nums = new();
nums.Add(1.5); nums.Add(2.5);
Console.WriteLine(nums[0] + nums[1]);"#,
        ["4"]
    };

    target_new_dictionary_char_int_inferred => {
        r#"System.Collections.Generic.Dictionary<char, int> map = new();
map['A'] = 1;
Console.WriteLine(map['A']);"#,
        ["1"]
    };

    target_new_list_with_object_initializer_inferred => {
        r#"System.Collections.Generic.List<int> values = new() { 1, 2, 3 };
Console.WriteLine(values[2]);"#,
        ["3"]
    };

    target_new_dictionary_with_object_initializer_inferred => {
        r#"System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 };
Console.WriteLine(map["b"]);"#,
        ["2"]
    };

    target_new_custom_class_inferred_from_local => {
        r#"class Widget { public int Id = 0; }
Widget w = new();
w.Id = 9;
Console.WriteLine(w.Id);"#,
        ["9"]
    };

    target_new_custom_class_as_field_initializer => {
        r#"class Holder { public System.Collections.Generic.List<int> items = new(); }
var h = new Holder();
h.items.Add(6);
Console.WriteLine(h.items[0]);"#,
        ["6"]
    };

    target_new_custom_class_returned_from_method => {
        r#"class Node { public int Value; }
Node Make() { Node n = new(); n.Value = 12; return n; }
Console.WriteLine(Make().Value);"#,
        ["12"]
    };

    target_new_list_passed_as_argument_inferred => {
        r#"int Sum(System.Collections.Generic.List<int> xs) { int s = 0; foreach (var x in xs) s += x; return s; }
System.Collections.Generic.List<int> data = new() { 1, 2, 3 };
Console.WriteLine(Sum(data));"#,
        ["6"]
    };

    target_new_nested_list_inferred => {
        r#"System.Collections.Generic.List<System.Collections.Generic.List<int>> grid = new();
System.Collections.Generic.List<int> row = new() { 1, 2 };
grid.Add(row);
Console.WriteLine(grid[0][1]);"#,
        ["2"]
    };

    target_new_list_long_inferred => {
        r#"System.Collections.Generic.List<long> longs = new();
longs.Add(9000000000L);
Console.WriteLine(longs[0]);"#,
        ["9000000000"]
    };

    target_new_list_decimal_inferred => {
        r#"System.Collections.Generic.List<decimal> prices = new();
prices.Add(9.99m);
Console.WriteLine(prices[0]);"#,
        ["9.99"]
    };

    target_new_hashset_string_inferred => {
        r#"System.Collections.Generic.HashSet<string> tags = new();
tags.Add("a"); tags.Add("b");
Console.WriteLine(tags.Contains("a")); Console.WriteLine(tags.Contains("c"));"#,
        ["True", "False"]
    };

    target_new_queue_string_inferred => {
        r#"System.Collections.Generic.Queue<string> q = new();
q.Enqueue("first"); q.Enqueue("second");
Console.WriteLine(q.Peek());"#,
        ["first"]
    };

    target_new_stack_string_inferred => {
        r#"System.Collections.Generic.Stack<string> s = new();
s.Push("base"); s.Push("top");
Console.WriteLine(s.Peek());"#,
        ["top"]
    };

    target_new_list_capacity_then_add => {
        r#"System.Collections.Generic.List<int> buf = new();
for (int i = 0; i < 4; i++) buf.Add(i);
Console.WriteLine(buf.Count);"#,
        ["4"]
    };

    target_new_dictionary_tryget_inferred => {
        r#"System.Collections.Generic.Dictionary<string, int> map = new();
map["x"] = 5;
Console.WriteLine(map.TryGetValue("x", out int v)); Console.WriteLine(v);"#,
        ["True", "5"]
    };

    target_new_list_of_custom_type_inferred => {
        r#"class Item { public string Name = ""; }
System.Collections.Generic.List<Item> items = new();
items.Add(new Item { Name = "tool" });
Console.WriteLine(items[0].Name);"#,
        ["tool"]
    };

    target_delegate_func_int_to_int_from_method_group => {
        r#"static int Double(int n) => n * 2;
System.Func<int, int> fn = Double;
Console.WriteLine(fn(6));"#,
        ["12"]
    };

    target_delegate_action_from_void_method_group => {
        r#"int total = 0;
void Bump() { total++; }
System.Action bump = Bump;
bump(); bump();
Console.WriteLine(total);"#,
        ["2"]
    };

    target_delegate_func_two_args_from_method_group => {
        r#"static int Add(int a, int b) => a + b;
System.Func<int, int, int> add = Add;
Console.WriteLine(add(3, 4));"#,
        ["7"]
    };

    target_delegate_predicate_from_method_group => {
        r#"static bool IsEven(int n) => n % 2 == 0;
System.Predicate<int> even = IsEven;
Console.WriteLine(even(4)); Console.WriteLine(even(3));"#,
        ["True", "False"]
    };

    target_delegate_func_string_to_int_from_method_group => {
        r#"static int Len(string s) => s.Length;
System.Func<string, int> measure = Len;
Console.WriteLine(measure("hello"));"#,
        ["5"]
    };

    target_delegate_action_string_from_method_group => {
        r#"string last = "";
void Capture(string s) { last = s; }
System.Action<string> store = Capture;
store("saved");
Console.WriteLine(last);"#,
        ["saved"]
    };

    target_delegate_func_inferred_from_lambda_assignment => {
        r#"System.Func<int, int> triple = x => x * 3;
Console.WriteLine(triple(4));"#,
        ["12"]
    };

    target_delegate_action_inferred_from_lambda => {
        r#"int count = 0;
System.Action tick = () => count++;
tick(); tick();
Console.WriteLine(count);"#,
        ["2"]
    };

    target_delegate_func_two_params_inferred_lambda => {
        r#"System.Func<int, int, int> mul = (a, b) => a * b;
Console.WriteLine(mul(3, 5));"#,
        ["15"]
    };

    target_delegate_func_returns_bool_inferred => {
        r#"System.Func<int, bool> positive = n => n > 0;
Console.WriteLine(positive(1)); Console.WriteLine(positive(-1));"#,
        ["True", "False"]
    };

    target_delegate_action_with_closure_capture => {
        r#"int sum = 0;
System.Action<int> add = n => sum += n;
add(3); add(4);
Console.WriteLine(sum);"#,
        ["7"]
    };

    target_delegate_instance_method_group_to_func => {
        r#"class Scale { public int factor = 2; public int Apply(int n) => n * factor; }
System.Func<int, int> fn = new Scale().Apply;
Console.WriteLine(fn(5));"#,
        ["10"]
    };

    target_delegate_instance_method_group_to_action => {
        r#"class Logger { public string last = ""; public void Save(string msg) => last = msg; }
var log = new Logger();
System.Action<string> write = log.Save;
write("note");
Console.WriteLine(log.last);"#,
        ["note"]
    };

    target_delegate_static_method_group_to_action => {
        r#"static int hits = 0;
static void Hit() { hits++; }
System.Action strike = Hit;
strike(); strike();
Console.WriteLine(hits);"#,
        ["2"]
    };

    target_delegate_func_chained_invocation => {
        r#"System.Func<int, int> inc = x => x + 1;
System.Func<int, int> twice = x => inc(inc(x));
Console.WriteLine(twice(3));"#,
        ["5"]
    };

    target_delegate_func_nullable_int_inferred => {
        r#"System.Func<int?, int> orZero = n => n ?? 0;
Console.WriteLine(orZero(null)); Console.WriteLine(orZero(7));"#,
        ["0", "7"]
    };

    target_delegate_comparison_func_inferred => {
        r#"System.Func<int, int, bool> less = (a, b) => a < b;
Console.WriteLine(less(2, 5)); Console.WriteLine(less(9, 1));"#,
        ["True", "False"]
    };

    target_new_and_delegate_list_foreach_with_action => {
        r#"System.Collections.Generic.List<int> nums = new() { 1, 2, 3 };
int sum = 0;
System.Action<int> acc = n => sum += n;
foreach (var n in nums) acc(n);
Console.WriteLine(sum);"#,
        ["6"]
    };

    target_new_dictionary_values_to_list_inferred => {
        r#"System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 };
System.Collections.Generic.List<int> values = new();
foreach (var kv in map) values.Add(kv.Value);
Console.WriteLine(values[0] + values[1]);"#,
        ["3"]
    };

    target_delegate_func_string_format_inferred => {
        r#"System.Func<string, string, string> join = (a, b) => a + "-" + b;
Console.WriteLine(join("x", "y"));"#,
        ["x-y"]
    };

    target_new_list_in_conditional_branch => {
        r#"System.Collections.Generic.List<int> pick(bool flag) {
    System.Collections.Generic.List<int> a = new() { 1 };
    System.Collections.Generic.List<int> b = new() { 2 };
    return flag ? a : b;
}
Console.WriteLine(pick(false)[0]);"#,
        ["2"]
    };

    target_delegate_func_returning_delegate_inferred => {
        r#"System.Func<int, System.Func<int, int>> scale = factor => n => n * factor;
var triple = scale(3);
Console.WriteLine(triple(4));"#,
        ["12"]
    };

    target_new_sorted_set_inferred_if_available => {
        r#"System.Collections.Generic.SortedSet<int> ordered = new();
ordered.Add(3); ordered.Add(1); ordered.Add(2);
foreach (var n in ordered) Console.WriteLine(n);"#,
        ["1", "2", "3"]
    };

    target_delegate_action_no_args_from_local_function => {
        r#"int n = 0;
void Reset() { n = 0; }
System.Action reset = Reset;
n = 5; reset();
Console.WriteLine(n);"#,
        ["0"]
    };

    target_new_list_clear_and_reuse_inferred => {
        r#"System.Collections.Generic.List<int> buf = new() { 1, 2 };
buf.Clear();
buf.Add(9);
Console.WriteLine(buf[0]);"#,
        ["9"]
    };

    target_delegate_func_char_to_int_inferred => {
        r#"System.Func<char, int> code = c => (int)c;
Console.WriteLine(code('A'));"#,
        ["65"]
    };
}
