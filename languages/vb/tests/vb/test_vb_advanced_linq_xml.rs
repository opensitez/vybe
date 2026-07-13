use super::helpers::run_vb;

#[test]
fn linq_aggregate_into() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4, 5}
        Dim total = Aggregate n In nums Into Sum()
        Console.WriteLine(total)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn linq_group_join() {
    let out = run_vb(
        r#"
Imports System.Linq
Imports System.Collections.Generic

Class Dept
    Public Id As Integer
    Public Name As String
End Class

Class Emp
    Public DeptId As Integer
    Public Name As String
End Class

Module M
    Sub Main()
        Dim depts = {New Dept With {.Id = 1, .Name = "IT"}}
        Dim emps = {New Emp With {.DeptId = 1, .Name = "Alice"}, New Emp With {.DeptId = 1, .Name = "Bob"}}
        
        Dim query = From d In depts
                    Group Join e In emps On d.Id Equals e.DeptId Into DeptEmps = Group
                    Select d.Name, Count = DeptEmps.Count()
                    
        For Each q In query
            Console.WriteLine(q.Name & "-" & q.Count)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["IT-2"]);
}

#[test]
fn linq_let_multiple() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        Dim query = From n In nums
                    Let x = n * 2, y = n * 3
                    Select x + y
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "10", "15"]); // 1*2 + 1*3 = 5, 2*2 + 2*3 = 10, etc.
}

#[test]
fn linq_join_multiple_keys() {
    let out = run_vb(
        r#"
Imports System.Linq

Class Item
    Public K1 As Integer
    Public K2 As Integer
    Public Val As String
End Class

Module M
    Sub Main()
        Dim arr1 = {New Item With {.K1 = 1, .K2 = 2, .Val = "A"}}
        Dim arr2 = {New Item With {.K1 = 1, .K2 = 2, .Val = "B"}}
        
        Dim query = From a In arr1
                    Join b In arr2 On a.K1 Equals b.K1 And a.K2 Equals b.K2
                    Select a.Val & b.Val
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn linq_skip_take_chain() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4, 5, 6, 7}
        Dim query = From n In nums
                    Skip 2
                    Take 3
                    Select n
                    
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "4", "5"]);
}

#[test]
fn linq_from_multiple() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim arr1 = {1, 2}
        Dim arr2 = {"A", "B"}
        
        Dim query = From a In arr1, b In arr2
                    Select a & b
                    
        For Each q In query
            Console.WriteLine(q)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1A", "1B", "2A", "2B"]);
}

#[test]
fn xml_literal_with_cdata_expr() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = "Test"
        ' Expressions inside CDATA are not evaluated in VB, they are literal text
        Dim xml = <Root><![CDATA[ <%= val %> ]]></Root>
        Console.WriteLine(xml.Value.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["<%= val %>"]);
}

#[test]
fn xml_literal_with_comment_expr() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = "Test"
        ' Expressions inside XML comments are also literal
        Dim xml = <Root><!-- <%= val %> --></Root>
        Console.WriteLine(xml.FirstNode.ToString().Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["<!-- <%= val %> -->"]);
}

#[test]
fn xml_axis_descendant_indexer() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <Root>
                      <Node>
                          <Child>1</Child>
                      </Node>
                      <Node>
                          <Child>2</Child>
                      </Node>
                  </Root>
                  
        Console.WriteLine(xml...<Child>(1).Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn xml_namespace_imports_alias() {
    let out = run_vb(
        r#"
Imports <xmlns:ns="http://test.com/ns">

Module M
    Sub Main()
        Dim xml = <ns:Root><ns:Child>Val</ns:Child></ns:Root>
        Console.WriteLine(xml.Name.NamespaceName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["http://test.com/ns"]);
}

#[test]
fn xml_literal_empty_element() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <Root />
        Console.WriteLine(xml.IsEmpty)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn xml_literal_document() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim doc = <?xml version="1.0"?><Root/>
        Console.WriteLine(doc.Root.Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Root"]);
}

#[test]
fn linq_distinct_with_comparer() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim strings = {"A", "a", "B"}
        ' Default distinct is case sensitive
        Dim q1 = strings.Distinct().Count()
        ' With Case Insensitive comparer
        Dim q2 = strings.Distinct(System.StringComparer.OrdinalIgnoreCase).Count()
        
        Console.WriteLine(q1)
        Console.WriteLine(q2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn linq_union_intersect_except() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim a = {1, 2, 3}
        Dim b = {3, 4, 5}
        
        Dim un = a.Union(b).Count()
        Dim int = a.Intersect(b).Count()
        Dim exc = a.Except(b).Count()
        
        Console.WriteLine(un & "-" & int & "-" & exc)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5-1-2"]);
}

#[test]
fn linq_any_all() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {2, 4, 6}
        Console.WriteLine(nums.Any(Function(n) n = 4))
        Console.WriteLine(nums.All(Function(n) n Mod 2 = 0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn linq_first_firstordefault() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums As Integer() = {}
        Console.WriteLine(nums.FirstOrDefault())
        
        Try
            nums.First()
        Catch
            Console.WriteLine("Error")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "Error"]);
}

#[test]
fn linq_single_singleordefault() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1}
        Console.WriteLine(nums.Single())
        
        Dim nums2 = {1, 2}
        Try
            nums2.SingleOrDefault()
        Catch
            Console.WriteLine("Error")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "Error"]);
}

#[test]
fn linq_last_lastordefault() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        Console.WriteLine(nums.Last())
        
        Dim empty As Integer() = {}
        Console.WriteLine(empty.LastOrDefault())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn linq_min_max_average() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {10, 20, 30}
        Console.WriteLine(nums.Min())
        Console.WriteLine(nums.Max())
        Console.WriteLine(nums.Average())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "30", "20"]);
}

#[test]
fn linq_count_longcount() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4}
        Console.WriteLine(nums.Count(Function(n) n > 2))
        Console.WriteLine(nums.LongCount())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn linq_sum() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim strings = {"A", "BB", "CCC"}
        Dim totalLen = strings.Sum(Function(s) s.Length)
        Console.WriteLine(totalLen)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn xml_create_element_dynamic() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim el As New XElement("Node", "Value")
        Console.WriteLine(el.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["<Node>Value</Node>"]);
}

#[test]
fn xml_create_attribute_dynamic() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim attr As New XAttribute("Id", "123")
        Dim el As New XElement("Node", attr)
        Console.WriteLine(el.Attribute("Id").Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn xml_add_remove_nodes() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim el = <Root><A/></Root>
        el.Add(<B/>)
        Console.WriteLine(el.Elements().Count())
        
        el.Element("A").Remove()
        Console.WriteLine(el.Elements().Count())
        Console.WriteLine(el.Elements().First().Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "1", "B"]);
}

#[test]
fn xml_replace_nodes() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim el = <Root><A/></Root>
        el.ReplaceNodes(<B/>)
        Console.WriteLine(el.Elements().First().Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn xml_set_attribute_value() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim el = <Root Id="1"/>
        el.SetAttributeValue("Id", "2")
        el.SetAttributeValue("NewAttr", "3")
        
        Console.WriteLine(el.Attribute("Id").Value)
        Console.WriteLine(el.Attribute("NewAttr").Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn linq_query_syntax_vs_method() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        ' Mix query syntax with method call
        Dim list = (From n In nums Select n * 2).ToList()
        
        Console.WriteLine(list.Count)
        Console.WriteLine(list(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn linq_deferred_execution() {
    let out = run_vb(
        r#"
Imports System.Linq
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        
        Dim query = From n In list Select n
        
        list.Add(3)
        
        ' Query is evaluated here, should see 3
        Console.WriteLine(query.Count())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn linq_let_multiple_variables() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim strings = {"Apple", "Banana"}
        
        Dim query = From s In strings
                    Let first = s(0)
                    Let len = s.Length
                    Select first & len
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A5", "B6"]);
}

#[test]
fn xml_axis_extension_method() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq
Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()>
    Public Function GetNames(elements As IEnumerable(Of XElement)) As IEnumerable(Of String)
        Return elements.Select(Function(e) e.Name.LocalName)
    End Function
End Module

Module M
    Sub Main()
        Dim xml = <Root><A/><B/></Root>
        Dim names = xml.Elements().GetNames()
        
        For Each name In names
            Console.WriteLine(name)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "B"]);
}
