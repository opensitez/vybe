use super::helpers::run_vb;

// Dictionary initializers
#[test] fn dict_init_basic() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim d As New Dictionary(Of Integer, String) From {{1, "A"}, {2, "B"}}: Console.WriteLine(d(2)): End Sub: End Module"#), vec!["B"]); }
#[test] fn dict_init_mixed() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim d As New Dictionary(Of Object, Object) From {{"A", 1}, {2, "B"}}: Console.WriteLine(d("A")): End Sub: End Module"#), vec!["1"]); }

// Collection initializers
#[test] fn list_init() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim l As New List(Of Integer) From {1, 2, 3}: Console.WriteLine(l.Count): End Sub: End Module"#), vec!["3"]); }
#[test] fn list_init_custom() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Class C: Implements IEnumerable: Public Sub Add(x As Integer): Console.WriteLine(x): End Sub: Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator: Return Nothing: End Function: End Class: Module M: Sub Main(): Dim c As New C From {10, 20}: End Sub: End Module"#), vec!["10", "20"]); }

// LINQ edge cases
#[test] fn linq_where_chain() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {1, 2, 3, 4}: Dim q = From x In n Where x > 1 Where x < 4 Select x: Console.WriteLine(q.Count()): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_order_multiple() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {New With {.A = 1, .B = 2}, New With {.A = 1, .B = 1}}: Dim q = From x In n Order By x.A, x.B Descending Select x.B: Console.WriteLine(q.First()): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_group_by_single() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {1, 2, 2, 3}: Dim q = From x In n Group By x Into Group Select x, Count = Group.Count(): Console.WriteLine(q.Where(Function(a) a.x = 2).First().Count): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_group_by_expression() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {1, 2, 3, 4}: Dim q = From x In n Group By IsEven = (x Mod 2 = 0) Into Group Select IsEven, Total = Group.Sum(): Console.WriteLine(q.Where(Function(a) a.IsEven).First().Total): End Sub: End Module"#), vec!["6"]); }
#[test] fn linq_group_by_multiple_keys() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {New With {.A = 1, .B = 1, .V = 10}, New With {.A = 1, .B = 1, .V = 20}}: Dim q = From x In n Group By x.A, x.B Into Group Select A, B, Total = Group.Sum(Function(g) g.V): Console.WriteLine(q.First().Total): End Sub: End Module"#), vec!["30"]); }
#[test] fn linq_join_simple() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}: Dim b = {2, 3}: Dim q = From x In a Join y In b On x Equals y Select x: Console.WriteLine(q.First()): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_group_join_simple() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}: Dim b = {2, 2, 3}: Dim q = From x In a Group Join y In b On x Equals y Into Matches = Group Select x, Count = Matches.Count(): Console.WriteLine(q.Where(Function(i) i.x = 2).First().Count): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_select_anonymous() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {1}: Dim q = From x In n Select New With {.V = x * 2}: Console.WriteLine(q.First().V): End Sub: End Module"#), vec!["2"]); }

// More Arrays & Collections
#[test] fn array_clone() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {1, 2}: Dim b = CType(a.Clone(), Integer()): b(0) = 10: Console.WriteLine(a(0)): End Sub: End Module"#), vec!["1"]); }
#[test] fn array_copyto() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {1, 2}: Dim b(1) As Integer: a.CopyTo(b, 0): Console.WriteLine(b(1)): End Sub: End Module"#), vec!["2"]); }
#[test] fn array_clear() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {1, 2}: Array.Clear(a, 0, 2): Console.WriteLine(a(0)): End Sub: End Module"#), vec!["0"]); }
#[test] fn array_reverse() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {1, 2, 3}: Array.Reverse(a): Console.WriteLine(a(0)): End Sub: End Module"#), vec!["3"]); }
#[test] fn array_sort() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {3, 1, 2}: Array.Sort(a): Console.WriteLine(a(0)): End Sub: End Module"#), vec!["1"]); }
#[test] fn array_indexof() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {10, 20, 30}: Console.WriteLine(Array.IndexOf(a, 20)): End Sub: End Module"#), vec!["1"]); }
#[test] fn array_binarysearch() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a = {10, 20, 30}: Console.WriteLine(Array.BinarySearch(a, 20)): End Sub: End Module"#), vec!["1"]); }

// Collection builtins (Microsoft.VisualBasic.Collection)
#[test] fn vb_collection_add() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add("A"): Console.WriteLine(c(1)): End Sub: End Module"#), vec!["A"]); }
#[test] fn vb_collection_add_key() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add("A", "Key1"): Console.WriteLine(c("Key1")): End Sub: End Module"#), vec!["A"]); }
#[test] fn vb_collection_count() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add(1): c.Add(2): Console.WriteLine(c.Count): End Sub: End Module"#), vec!["2"]); }
#[test] fn vb_collection_remove() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add("A", "K"): c.Remove("K"): Console.WriteLine(c.Count): End Sub: End Module"#), vec!["0"]); }
#[test] fn vb_collection_contains() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add("A", "K"): Console.WriteLine(c.Contains("K")): End Sub: End Module"#), vec!["True"]); }

// LINQ Let, Distinct, Union, Intersect, Except
#[test] fn linq_let_multiple_use() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim n = {1}: Dim q = From x In n Let y = x + 1, z = y + 1 Select x + y + z: Console.WriteLine(q.First()): End Sub: End Module"#), vec!["6"]); }
#[test] fn linq_distinct_struct() { assert_eq!(run_vb(r#"Imports System.Linq: Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim n = {New S With {.V = 1}, New S With {.V = 1}}: Console.WriteLine(n.Distinct().Count()): End Sub: End Module"#), vec!["1"]); }
#[test] fn linq_union_struct() { assert_eq!(run_vb(r#"Imports System.Linq: Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim a = {New S With {.V = 1}}: Dim b = {New S With {.V = 1}}: Console.WriteLine(a.Union(b).Count()): End Sub: End Module"#), vec!["1"]); }
#[test] fn linq_intersect_struct() { assert_eq!(run_vb(r#"Imports System.Linq: Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim a = {New S With {.V = 1}}: Dim b = {New S With {.V = 1}}: Console.WriteLine(a.Intersect(b).Count()): End Sub: End Module"#), vec!["1"]); }
#[test] fn linq_except_struct() { assert_eq!(run_vb(r#"Imports System.Linq: Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim a = {New S With {.V = 1}}: Dim b = {New S With {.V = 1}}: Console.WriteLine(a.Except(b).Count()): End Sub: End Module"#), vec!["0"]); }

// LINQ Cast, OfType
#[test] fn linq_cast() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a As Object() = {1, 2}: Console.WriteLine(a.Cast(Of Integer)().Sum()): End Sub: End Module"#), vec!["3"]); }
#[test] fn linq_oftype() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a As Object() = {1, "Two", 3}: Console.WriteLine(a.OfType(Of Integer)().Sum()): End Sub: End Module"#), vec!["4"]); }

// LINQ Any, All with empty
#[test] fn linq_any_empty() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a As Integer() = {}: Console.WriteLine(a.Any()): End Sub: End Module"#), vec!["False"]); }
#[test] fn linq_all_empty() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a As Integer() = {}: Console.WriteLine(a.All(Function(x) x > 0)): End Sub: End Module"#), vec!["True"]); }

// LINQ ToArray, ToList, ToDictionary, ToLookup
#[test] fn linq_toarray() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}.Where(Function(x) x > 1).ToArray(): Console.WriteLine(a.Length): End Sub: End Module"#), vec!["1"]); }
#[test] fn linq_tolist() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim l = {1, 2}.ToList(): Console.WriteLine(l.Count): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_todictionary() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim d = {1, 2}.ToDictionary(Function(x) x.ToString()): Console.WriteLine(d("2")): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_tolookup() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim l = {1, 2, 2}.ToLookup(Function(x) x.ToString()): Console.WriteLine(l("2").Count()): End Sub: End Module"#), vec!["2"]); }

// HashSets and Queues
#[test] fn hashset_init() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim h As New HashSet(Of Integer) From {1, 1, 2}: Console.WriteLine(h.Count): End Sub: End Module"#), vec!["2"]); }
#[test] fn queue_enqueue_dequeue() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim q As New Queue(Of Integer)(): q.Enqueue(10): Console.WriteLine(q.Dequeue()): End Sub: End Module"#), vec!["10"]); }
#[test] fn stack_push_pop() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim s As New Stack(Of Integer)(): s.Push(20): Console.WriteLine(s.Pop()): End Sub: End Module"#), vec!["20"]); }
#[test] fn linkedlist_add() { assert_eq!(run_vb(r#"Imports System.Collections.Generic: Module M: Sub Main(): Dim l As New LinkedList(Of Integer)(): l.AddLast(1): l.AddLast(2): Console.WriteLine(l.First.Next.Value): End Sub: End Module"#), vec!["2"]); }

// More arrays
#[test] fn array_redim_2d() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a(,) As Integer: ReDim a(1, 1): a(1, 1) = 42: Console.WriteLine(a(1, 1)): End Sub: End Module"#), vec!["42"]); }
#[test] fn array_redim_preserve_2d() { assert_eq!(run_vb(r#"Module M: Sub Main(): Dim a(1, 1) As Integer: a(1, 1) = 42: ReDim Preserve a(1, 2): Console.WriteLine(a(1, 1)): End Sub: End Module"#), vec!["42"]); }

// Advanced Linq DefaultIfEmpty
#[test] fn linq_defaultifempty() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a As Integer() = {}: Dim q = a.DefaultIfEmpty(5): Console.WriteLine(q.First()): End Sub: End Module"#), vec!["5"]); }
#[test] fn linq_zip() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}: Dim b = {3, 4}: Dim q = a.Zip(b, Function(x, y) x + y): Console.WriteLine(q.Last()): End Sub: End Module"#), vec!["6"]); }
#[test] fn linq_concat() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1}: Dim b = {2}: Dim q = a.Concat(b): Console.WriteLine(q.Count()): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_reverse() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}: Dim q = a.Reverse(): Console.WriteLine(q.First()): End Sub: End Module"#), vec!["2"]); }
#[test] fn linq_sequenceequal() { assert_eq!(run_vb(r#"Imports System.Linq: Module M: Sub Main(): Dim a = {1, 2}: Dim b = {1, 2}: Console.WriteLine(a.SequenceEqual(b)): End Sub: End Module"#), vec!["True"]); }
