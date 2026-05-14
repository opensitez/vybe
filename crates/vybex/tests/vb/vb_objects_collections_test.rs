//! Comprehensive end-to-end tests for VB objects, collections, and arrays.
//!
//! Categories:
//!   A. Object creation and properties (10 tests)
//!   B. Collections -- List (10 tests)
//!   C. Collections -- Dictionary (8 tests)
//!   D. Collections -- Queue and Stack (6 tests)
//!   E. Array operations (10 tests)
//!   F. Passing objects between functions (8 tests)
//!   G. Namespace object creation (8 tests)

use super::helpers::run_vb;

// ============================================================
// A. Object creation and properties (10 tests)
// ============================================================

#[test]
fn a01_create_class_set_and_read_property() {
    let out = run_vb(r#"
Class Dog
    Public Name As String
End Class
Dim d As New Dog()
d.Name = "Rex"
Console.WriteLine(d.Name)
"#);
    assert_eq!(out, vec!["Rex"]);
}

#[test]
fn a02_object_with_multiple_fields() {
    let out = run_vb(r#"
Class Person
    Public Name As String
    Public Age As Integer
    Public City As String
End Class
Dim p As New Person()
p.Name = "Alice"
p.Age = 30
p.City = "Paris"
Console.WriteLine(p.Name)
Console.WriteLine(p.Age)
Console.WriteLine(p.City)
"#);
    assert_eq!(out, vec!["Alice", "30", "Paris"]);
}

#[test]
fn a03_nested_objects() {
    let out = run_vb(r#"
Class Address
    Public City As String
End Class
Class Person
    Public Name As String
    Public Addr As Address
End Class
Dim a As New Address()
a.City = "London"
Dim p As New Person()
p.Name = "Bob"
p.Addr = a
Console.WriteLine(p.Addr.City)
"#);
    assert_eq!(out, vec!["London"]);
}

#[test]
fn a04_object_passed_to_function_modified() {
    let out = run_vb(r#"
Class Counter
    Public Value As Integer
End Class
Sub Increment(c As Counter)
    c.Value = c.Value + 1
End Sub
Dim c As New Counter()
c.Value = 10
Increment(c)
Console.WriteLine(c.Value)
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn a05_object_returned_from_function() {
    let out = run_vb(r#"
Class Item
    Public Label As String
End Class
Function MakeItem(lbl As String) As Item
    Dim it As New Item()
    it.Label = lbl
    Return it
End Function
Dim x As Item = MakeItem("hello")
Console.WriteLine(x.Label)
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn a06_object_stored_in_array() {
    let out = run_vb(r#"
Class Box
    Public Size As Integer
End Class
Dim arr(2) As Box
Dim b As New Box()
b.Size = 42
arr(0) = b
Console.WriteLine(arr(0).Size)
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn a07_two_instances_independent_state() {
    let out = run_vb(r#"
Class Widget
    Public Color As String
End Class
Dim w1 As New Widget()
Dim w2 As New Widget()
w1.Color = "red"
w2.Color = "blue"
Console.WriteLine(w1.Color)
Console.WriteLine(w2.Color)
"#);
    assert_eq!(out, vec!["red", "blue"]);
}

#[test]
fn a08_object_field_initialized_to_nothing_then_set() {
    let out = run_vb(r#"
Class Holder
    Public Item As Object
End Class
Dim h As New Holder()
Console.WriteLine(IsNothing(h.Item))
h.Item = "something"
Console.WriteLine(h.Item)
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true", "something"]));
}

#[test]
fn a09_object_field_set_to_another_object() {
    let out = run_vb(r#"
Class Inner
    Public Val As Integer
End Class
Class Outer
    Public Child As Inner
End Class
Dim i As New Inner()
i.Val = 99
Dim o As New Outer()
o.Child = i
Console.WriteLine(o.Child.Val)
"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn a10_object_identity_same_instance() {
    let out = run_vb(r#"
Class Bag
    Public Count As Integer
End Class
Dim a As New Bag()
a.Count = 5
Dim b As Bag = a
b.Count = 10
Console.WriteLine(a.Count)
"#);
    assert_eq!(out, vec!["10"]);
}

// ============================================================
// B. Collections -- List (10 tests)
// ============================================================

#[test]
fn b11_list_add_and_count() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("hello")
list.Add("world")
Console.WriteLine(list.Count)
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn b12_list_for_each_order() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
For Each item As String In list
    Console.WriteLine(item)
Next
"#);
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn b13_list_item_by_index() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("x")
list.Add("y")
list.Add("z")
Console.WriteLine(list.Item(1))
"#);
    assert_eq!(out, vec!["y"]);
}

#[test]
fn b14_list_contains() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("apple")
list.Add("banana")
Console.WriteLine(list.Contains("apple"))
Console.WriteLine(list.Contains("cherry"))
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true", "false"]));
}

#[test]
fn b15_list_remove() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
list.Remove("b")
Console.WriteLine(list.Count)
For Each item As String In list
    Console.WriteLine(item)
Next
"#);
    assert_eq!(out, vec!["2", "a", "c"]);
}

#[test]
fn b16_list_clear() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("x")
list.Add("y")
list.Clear()
Console.WriteLine(list.Count)
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn b17_list_indexof() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("one")
list.Add("two")
list.Add("three")
Console.WriteLine(list.IndexOf("two"))
Console.WriteLine(list.IndexOf("four"))
"#);
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn b18_list_of_numbers_sum() {
    let out = run_vb(r#"
Dim list As New List(Of Integer)
list.Add(10)
list.Add(20)
list.Add(30)
Dim total As Integer = 0
For Each n As Integer In list
    total = total + n
Next
Console.WriteLine(total)
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn b19_list_add_objects() {
    let out = run_vb(r#"
Class Cat
    Public Name As String
End Class
Dim list As New List(Of Cat)
Dim c1 As New Cat()
c1.Name = "Whiskers"
Dim c2 As New Cat()
c2.Name = "Mittens"
list.Add(c1)
list.Add(c2)
Console.WriteLine(list.Item(0).Name)
Console.WriteLine(list.Item(1).Name)
"#);
    assert_eq!(out, vec!["Whiskers", "Mittens"]);
}

#[test]
fn b20_list_count_after_add_remove() {
    let out = run_vb(r#"
Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
list.Add("d")
list.Remove("b")
list.Remove("d")
Console.WriteLine(list.Count)
"#);
    assert_eq!(out, vec!["2"]);
}

// ============================================================
// C. Collections -- Dictionary (8 tests)
// ============================================================

#[test]
fn c21_dictionary_add_and_access() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("name", "Alice")
Console.WriteLine(dict.Item("name"))
"#);
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn c22_dictionary_containskey() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("key1", "val1")
Console.WriteLine(dict.ContainsKey("key1"))
Console.WriteLine(dict.ContainsKey("key2"))
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["true", "false"]));
}

#[test]
fn c23_dictionary_remove() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("a", "1")
dict.Add("b", "2")
dict.Remove("a")
Console.WriteLine(dict.ContainsKey("a"))
Console.WriteLine(dict.ContainsKey("b"))
"#);
    assert_eq!(out, super::helpers::dotnet_expected_lines(&["false", "true"]));
}

#[test]
fn c24_dictionary_count() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("x", "1")
dict.Add("y", "2")
dict.Add("z", "3")
Console.WriteLine(dict.Count)
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn c25_dictionary_overwrite_value() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("key", "old")
dict.Add("key", "new")
Console.WriteLine(dict.Item("key"))
"#);
    assert_eq!(out, vec!["new"]);
}

#[test]
fn c26_dictionary_iterate_keys() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, String)
dict.Add("alpha", "1")
dict.Add("beta", "2")
Dim keys = dict.Keys()
For Each k As String In keys
    Console.WriteLine(k)
Next
"#);
    // Dictionary key order is not guaranteed, so sort
    let mut sorted = out.clone();
    sorted.sort();
    assert_eq!(sorted, vec!["alpha", "beta"]);
}

#[test]
fn c27_dictionary_iterate_values() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, Integer)
dict.Add("a", 10)
dict.Add("b", 20)
Dim vals = dict.Values()
Dim total As Integer = 0
For Each v As Integer In vals
    total = total + v
Next
Console.WriteLine(total)
"#);
    assert_eq!(out, vec!["30"]);
}


#[test]
fn c28_dictionary_integer_values() {
    let out = run_vb(r#"
Dim dict As New Dictionary(Of String, Integer)
dict.Add("score", 100)
dict.Add("bonus", 50)
Dim s As Integer = dict.Item("score") + dict.Item("bonus")
Console.WriteLine(s)
"#);
    assert_eq!(out, vec!["150"]);
}

// ============================================================
// D. Collections -- Queue and Stack (6 tests)
// ============================================================

#[test]
fn d29_queue_fifo_order() {
    let out = run_vb(r#"
Dim q As New Queue(Of String)
q.Enqueue("first")
q.Enqueue("second")
q.Enqueue("third")
Console.WriteLine(q.Dequeue())
Console.WriteLine(q.Dequeue())
Console.WriteLine(q.Dequeue())
"#);
    assert_eq!(out, vec!["first", "second", "third"]);
}

#[test]
fn d30_queue_count_and_peek() {
    let out = run_vb(r#"
Dim q As New Queue(Of String)
q.Enqueue("a")
q.Enqueue("b")
Console.WriteLine(q.Count)
Console.WriteLine(q.Peek())
Console.WriteLine(q.Count)
"#);
    assert_eq!(out, vec!["2", "a", "2"]);
}

#[test]
fn d31_queue_process_until_empty() {
    let out = run_vb(r#"
Dim q As New Queue(Of Integer)
q.Enqueue(1)
q.Enqueue(2)
q.Enqueue(3)
Dim total As Integer = 0
Do While q.Count > 0
    total = total + q.Dequeue()
Loop
Console.WriteLine(total)
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn d32_stack_lifo_order() {
    let out = run_vb(r#"
Dim s As New Stack(Of String)
s.Push("first")
s.Push("second")
s.Push("third")
Console.WriteLine(s.Pop())
Console.WriteLine(s.Pop())
Console.WriteLine(s.Pop())
"#);
    assert_eq!(out, vec!["third", "second", "first"]);
}

#[test]
fn d33_stack_count_and_peek() {
    let out = run_vb(r#"
Dim s As New Stack(Of String)
s.Push("a")
s.Push("b")
Console.WriteLine(s.Count)
Console.WriteLine(s.Peek())
Console.WriteLine(s.Count)
"#);
    assert_eq!(out, vec!["2", "b", "2"]);
}

#[test]
fn d34_stack_process_until_empty() {
    let out = run_vb(r#"
Dim s As New Stack(Of Integer)
s.Push(10)
s.Push(20)
s.Push(30)
Dim total As Integer = 0
Do While s.Count > 0
    total = total + s.Pop()
Loop
Console.WriteLine(total)
"#);
    assert_eq!(out, vec!["60"]);
}

// ============================================================
// E. Array operations (10 tests)
// ============================================================

#[test]
fn e35_array_set_and_read_by_index() {
    let out = run_vb(r#"
Dim arr(5) As Integer
arr(0) = 10
arr(3) = 42
Console.WriteLine(arr(0))
Console.WriteLine(arr(3))
"#);
    assert_eq!(out, vec!["10", "42"]);
}

#[test]
fn e36_array_ubound() {
    let out = run_vb(r#"
Dim arr(5) As Integer
Console.WriteLine(UBound(arr))
"#);
    // Dim arr(5) creates 6 elements (indices 0-5), UBound = 5
    assert_eq!(out, vec!["5"]);
}

#[test]
fn e37_array_for_each() {
    // Dim arr(3) creates 4 elements (indices 0-3), only 3 assigned
    // Use arr(2) to match exactly the elements we set
    let out = run_vb(r#"
Dim arr(2) As Integer
arr(0) = 1
arr(1) = 2
arr(2) = 3
For Each n As Integer In arr
    Console.WriteLine(n)
Next
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn e38_array_redim() {
    let out = run_vb(r#"
Dim arr(3) As Integer
arr(0) = 99
ReDim arr(5)
Console.WriteLine(arr(0))
Console.WriteLine(UBound(arr))
"#);
    // ReDim without Preserve clears the array; Console.WriteLine(Nothing) prints blank.
    // Dim arr(5) → 6 elements, UBound = 5
    assert_eq!(out, vec!["", "5"]);
}

#[test]
fn e39_array_redim_preserve() {
    let out = run_vb(r#"
Dim arr(3) As Integer
arr(0) = 10
arr(1) = 20
arr(2) = 30
ReDim Preserve arr(5)
Console.WriteLine(arr(0))
Console.WriteLine(arr(1))
Console.WriteLine(arr(2))
Console.WriteLine(UBound(arr))
"#);
    // ReDim Preserve arr(5) → 6 elements (0-5), UBound = 5
    assert_eq!(out, vec!["10", "20", "30", "5"]);
}

#[test]
fn e40_array_of_strings() {
    let out = run_vb(r#"
Dim arr(3) As String
arr(0) = "hello"
arr(1) = "world"
Console.WriteLine(arr(0))
Console.WriteLine(arr(1))
"#);
    assert_eq!(out, vec!["hello", "world"]);
}

#[test]
fn e41_array_of_objects() {
    let out = run_vb(r#"
Class Pt
    Public X As Integer
    Public Y As Integer
End Class
Dim pts(3) As Pt
Dim p As New Pt()
p.X = 5
p.Y = 10
pts(0) = p
Console.WriteLine(pts(0).X)
Console.WriteLine(pts(0).Y)
"#);
    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn e42_array_passed_to_sub() {
    let out = run_vb(r#"
Sub SetFirst(a() As Integer, val As Integer)
    a(0) = val
End Sub
Dim arr(3) As Integer
SetFirst(arr, 77)
Console.WriteLine(arr(0))
"#);
    assert_eq!(out, vec!["77"]);
}

#[test]

fn e43_array_returned_from_function() {
    let out = run_vb(r#"
Function MakeArr() As Integer()
    Dim a(3) As Integer
    a(0) = 1
    a(1) = 2
    a(2) = 3
    Return a
End Function
Dim result() As Integer = MakeArr()
Console.WriteLine(result(0))
Console.WriteLine(result(1))
Console.WriteLine(result(2))
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]

fn e44_array_of_arrays() {
    let out = run_vb(r#"
Dim outer(2) As Object
Dim inner1(2) As Integer
inner1(0) = 10
inner1(1) = 20
Dim inner2(2) As Integer
inner2(0) = 30
inner2(1) = 40
outer(0) = inner1
outer(1) = inner2
Console.WriteLine(outer(0)(0))
Console.WriteLine(outer(1)(1))
"#);
    assert_eq!(out, vec!["10", "40"]);
}

// ============================================================
// F. Passing objects between functions (8 tests)
// ============================================================

#[test]
fn f45_sub_modifies_object_field() {
    let out = run_vb(r#"
Class Score
    Public Points As Integer
End Class
Sub AddPoints(s As Score, p As Integer)
    s.Points = s.Points + p
End Sub
Dim sc As New Score()
sc.Points = 100
AddPoints(sc, 50)
Console.WriteLine(sc.Points)
"#);
    assert_eq!(out, vec!["150"]);
}

#[test]
fn f46_function_returns_new_object() {
    let out = run_vb(r#"
Class Result
    Public Status As String
End Class
Function GetResult() As Result
    Dim r As New Result()
    r.Status = "OK"
    Return r
End Function
Dim res As Result = GetResult()
Console.WriteLine(res.Status)
"#);
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn f47_sub_takes_list_adds_items() {
    let out = run_vb(r#"
Sub FillList(lst As List)
    lst.Add("one")
    lst.Add("two")
End Sub
Dim myList As New List(Of String)
FillList(myList)
Console.WriteLine(myList.Count)
Console.WriteLine(myList.Item(0))
Console.WriteLine(myList.Item(1))
"#);
    assert_eq!(out, vec!["2", "one", "two"]);
}

#[test]
fn f48_function_swaps_field_between_objects() {
    let out = run_vb(r#"
Class Pair
    Public Val As String
End Class
Sub SwapVals(a As Pair, b As Pair)
    Dim tmp As String = a.Val
    a.Val = b.Val
    b.Val = tmp
End Sub
Dim p1 As New Pair()
p1.Val = "hello"
Dim p2 As New Pair()
p2.Val = "world"
SwapVals(p1, p2)
Console.WriteLine(p1.Val)
Console.WriteLine(p2.Val)
"#);
    assert_eq!(out, vec!["world", "hello"]);
}

#[test]
fn f49_object_with_method_passed_to_function() {
    let out = run_vb(r#"
Class Greeter
    Public Name As String
    Public Function Greet() As String
        Return "Hello, " & Name
    End Function
End Class
Function GetGreeting(g As Greeter) As String
    Return g.Greet()
End Function
Dim gr As New Greeter()
gr.Name = "World"
Console.WriteLine(GetGreeting(gr))
"#);
    assert_eq!(out, vec!["Hello, World"]);
}

#[test]
fn f50_recursive_function_with_object_accumulator() {
    let out = run_vb(r#"
Class Acc
    Public Total As Integer
End Class
Sub AddUp(a As Acc, n As Integer)
    If n <= 0 Then Return
    a.Total = a.Total + n
    AddUp(a, n - 1)
End Sub
Dim acc As New Acc()
acc.Total = 0
AddUp(acc, 5)
Console.WriteLine(acc.Total)
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn f51_chain_functions_pass_object() {
    let out = run_vb(r#"
Class Data
    Public Value As String
End Class
Function CreateData() As Data
    Dim d As New Data()
    d.Value = "start"
    Return d
End Function
Sub TransformData(d As Data)
    d.Value = d.Value & "-transformed"
End Sub
Sub FinalizeData(d As Data)
    d.Value = d.Value & "-done"
End Sub
Dim d As Data = CreateData()
TransformData(d)
FinalizeData(d)
Console.WriteLine(d.Value)
"#);
    assert_eq!(out, vec!["start-transformed-done"]);
}

#[test]

fn f52_object_in_collection_method_called() {
    let out = run_vb(r#"
Class Calc
    Public Base As Integer
    Public Function Double() As Integer
        Return Base * 2
    End Function
End Class
Dim list As New List(Of Calc)
Dim c As New Calc()
c.Base = 21
list.Add(c)
Dim retrieved As Calc = list.Item(0)
Console.WriteLine(retrieved.Double())
"#);
    assert_eq!(out, vec!["42"]);
}

// ============================================================
// G. Namespace object creation (8 tests)
// ============================================================

#[test]
fn g53_new_point_properties() {
    let out = run_vb(r#"
Dim p As New System.Drawing.Point(10, 20)
Console.WriteLine(p.x)
Console.WriteLine(p.y)
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn g54_new_size_properties() {
    let out = run_vb(r#"
Dim s As New System.Drawing.Size(640, 480)
Console.WriteLine(s.width)
Console.WriteLine(s.height)
"#);
    assert_eq!(out, vec!["640", "480"]);
}

#[test]
fn g55_new_font_properties() {
    let out = run_vb(r#"
Dim f As New System.Drawing.Font("Arial", 12)
Console.WriteLine(f.name)
Console.WriteLine(f.size)
"#);
    assert_eq!(out, vec!["Arial", "12"]);
}

#[test]
fn g56_new_button_control_type() {
    let out = run_vb(r#"
Dim btn As New System.Windows.Forms.Button()
Console.WriteLine(btn.__control_type)
"#);
    assert_eq!(out, vec!["Button"]);
}

#[test]
fn g57_new_textbox_control_type() {
    let out = run_vb(r#"
Dim txt As New System.Windows.Forms.TextBox()
Console.WriteLine(txt.__control_type)
"#);
    assert_eq!(out, vec!["TextBox"]);
}

#[test]
fn g58_new_label_control_type() {
    let out = run_vb(r#"
Dim lbl As New System.Windows.Forms.Label()
Console.WriteLine(lbl.__control_type)
"#);
    assert_eq!(out, vec!["Label"]);
}

#[test]
fn g59_point_assigned_to_control_location() {
    let out = run_vb(r#"
Dim btn As New Button()
Dim pt As New Point(50, 100)
btn.Location = pt
Console.WriteLine(btn.Location.x)
Console.WriteLine(btn.Location.y)
"#);
    assert_eq!(out, vec!["50", "100"]);
}

#[test]
fn g60_size_assigned_to_control_size() {
    let out = run_vb(r#"
Dim btn As New Button()
Dim sz As New Size(200, 50)
btn.Size = sz
Console.WriteLine(btn.Size.width)
Console.WriteLine(btn.Size.height)
"#);
    assert_eq!(out, vec!["200", "50"]);
}
