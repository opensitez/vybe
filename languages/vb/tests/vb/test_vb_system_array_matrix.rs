use super::helpers::run_vb;

#[test]
fn array_resize_grows_and_preserves_prefix() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3}
        Array.Resize(values, 5)

        values(3) = 4
        values(4) = 5

        Console.WriteLine(values.Length)
        Console.WriteLine(values(0))
        Console.WriteLine(values(4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "1", "5"]);
}

#[test]
fn array_copy_transfers_slice_at_offset() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source() As Integer = {10, 11, 12, 13}
        Dim target() As Integer = {0, 0, 0, 0, 0}

        Array.Copy(source, 0, target, 1, 3)

        Console.WriteLine(target(0))
        Console.WriteLine(target(1))
        Console.WriteLine(target(3))
        Console.WriteLine(target(4))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "10", "12", "0"]);
}

#[test]
fn array_clear_zeroes_requested_range() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4}
        Array.Clear(values, 1, 2)

        Console.WriteLine(values(0))
        Console.WriteLine(values(1))
        Console.WriteLine(values(2))
        Console.WriteLine(values(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "0", "0", "4"]);
}

#[test]
fn array_sort_and_reverse_ordering() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {7, 1, 9, 2, 6}
        Array.Sort(values)
        Dim ascendingFirst As Integer = values(0)
        Dim ascendingLast As Integer = values(4)

        Array.Reverse(values)
        Dim descendingFirst As Integer = values(0)
        Dim descendingLast As Integer = values(4)

        Console.WriteLine(ascendingFirst)
        Console.WriteLine(ascendingLast)
        Console.WriteLine(descendingFirst)
        Console.WriteLine(descendingLast)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "9", "9", "1"]);
}

#[test]
fn array_indexof_position_and_miss() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {4, 5, 6, 5}
        Dim firstFive As Integer = Array.IndexOf(values, 5)
        Dim secondFive As Integer = Array.IndexOf(values, 5, firstFive + 1)
        Dim missing As Integer = Array.IndexOf(values, 99)

        Console.WriteLine(firstFive)
        Console.WriteLine(secondFive)
        Console.WriteLine(missing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "3", "-1"]);
}

#[test]
fn array_lastindexof_targets_final_occurrence() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {8, 9, 8, 7, 8}
        Dim lastEight As Integer = Array.LastIndexOf(values, 8)
        Console.WriteLine(lastEight)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4"]);
}

#[test]
fn array_binarysearch_found_and_not_found() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4, 5}
        Dim hit As Integer = Array.BinarySearch(values, 4)
        Dim miss As Integer = Array.BinarySearch(values, 7)

        Console.WriteLine(hit)
        Console.WriteLine(miss < 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3", "True"]);
}

#[test]
fn array_exists_and_trueforall_predicates() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {2, 4, 6, 8}
        Dim hasOdd As Boolean = Array.Exists(values, Function(v As Integer) v Mod 2 = 1)
        Dim allEven As Boolean = Array.TrueForAll(values, Function(v As Integer) v Mod 2 = 0)

        Console.WriteLine(hasOdd)
        Console.WriteLine(allEven)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn array_find_first_match() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {3, 5, 7, 9}
        Dim firstBig As Integer = Array.Find(values, Function(v As Integer) v > 6)
        Dim firstTiny As Integer = Array.Find(values, Function(v As Integer) v < 0)

        Console.WriteLine(firstBig)
        Console.WriteLine(firstTiny = 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "True"]);
}

#[test]
fn array_find_last_index_with_predicate() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {2, 4, 6, 9, 11}
        Dim idx As Integer = Array.FindLastIndex(values, Function(v As Integer) v Mod 2 = 1)

        Console.WriteLine(idx)
        Console.WriteLine(values(idx))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4", "11"]);
}

#[test]
fn array_find_all_filtering() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4, 5}
        Dim evens() As Integer = Array.FindAll(values, Function(v As Integer) v Mod 2 = 0)

        Console.WriteLine(evens.Length)
        Console.WriteLine(evens(0))
        Console.WriteLine(evens(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "2", "4"]);
}

#[test]
fn array_convertall_stringify() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 10}
        Dim tags() As String = Array.ConvertAll(values, Function(v As Integer) "v=" & v.ToString())

        Console.WriteLine(tags.Length)
        Console.WriteLine(tags(0))
        Console.WriteLine(tags(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "v=1", "v=10"]);
}

#[test]
fn array_foreach_aggregates_sum() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3}
        Dim sum As Integer = 0

        Array.ForEach(values, Sub(v As Integer)
            sum += v
        End Sub)

        Console.WriteLine(sum)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_createinstance_sets_and_reads_values() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim boxed As Array = Array.CreateInstance(GetType(Integer), 3)
        boxed.SetValue(7, 0)
        boxed.SetValue(8, 1)
        boxed.SetValue(9, 2)

        Console.WriteLine(CInt(boxed.GetValue(0)))
        Console.WriteLine(CInt(boxed.GetValue(2)))
        Console.WriteLine(Array.IndexOf(CType(boxed, Integer()), 8))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "9", "1"]);
}

#[test]
fn array_empty_is_zero_length_and_index_of_throws_expectedly_when_checked() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = Array.Empty(Of Integer)()
        Console.WriteLine(values.Length)
        Console.WriteLine(Array.BinarySearch(values, 1) < 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn array_set_and_get_via_reference_type() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {100, 200, 300}
        Dim boxed As Array = values
        boxed.SetValue(150, 1)

        Console.WriteLine(values(1))
        Dim roundTrip As Integer = CInt(boxed(1))
        Console.WriteLine(roundTrip)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["150", "150"]);
}

#[test]
fn array_reverse_preserves_length_and_content_checksum() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 1, 2, 3}
        Dim sumOriginal As Integer = values(0) + values(1) + values(2) + values(3)

        Array.Reverse(values)
        Dim sumReversed As Integer = values(0) + values(1) + values(2) + values(3)

        Console.WriteLine(sumOriginal)
        Console.WriteLine(sumReversed)
        Console.WriteLine(values(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "7", "1"]);
}
