use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ArraySegment & Memory Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_array_segment_creation_offset_count() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 3)
        Console.WriteLine(segment.Offset)
        Console.WriteLine(segment.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "3"]);
}

#[test]
fn test_vb_array_segment_element_access() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 3)
        Console.WriteLine(segment(0))
        Console.WriteLine(segment(2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20", "40"]);
}

#[test]
fn test_vb_array_segment_enumeration() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 3)
        Dim sum As Integer = 0
        For Each val In segment
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["90"]);
}

#[test]
fn test_vb_array_segment_slice_subsegment() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40, 50}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 4)
        Dim subSeg As ArraySegment(Of Integer) = segment.Slice(1, 2)
        Console.WriteLine(subSeg(0))
        Console.WriteLine(subSeg(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30", "40"]);
}

#[test]
fn test_vb_array_segment_underlying_array_mutation() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim segment As New ArraySegment(Of Integer)(numbers)
        segment(1) = 99
        Console.WriteLine(numbers(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}

#[test]
fn test_vb_array_segment_empty_property() {
    let src = r#"
Module Program
    Sub Main()
        Dim emptySeg As ArraySegment(Of Integer) = ArraySegment(Of Integer).Empty
        Console.WriteLine(emptySeg.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_array_segment_copy_to_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30, 40}
        Dim segment As New ArraySegment(Of Integer)(numbers, 1, 2)
        Dim dest(1) As Integer
        segment.CopyTo(dest)
        Console.WriteLine(String.Join(",", dest))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20,30"]);
}

#[test]
fn test_vb_array_segment_equals_operator() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim seg1 As New ArraySegment(Of Integer)(numbers, 0, 2)
        Dim seg2 As New ArraySegment(Of Integer)(numbers, 0, 2)
        Console.WriteLine(seg1 = seg2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_memory_copy_buffer() {
    let src = r#"
Module Program
    Sub Main()
        Dim srcArr As Byte() = {1, 2, 3, 4, 5}
        Dim dstArr(4) As Byte
        Buffer.BlockCopy(srcArr, 0, dstArr, 0, 5)
        Console.WriteLine(String.Join(",", dstArr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2,3,4,5"]);
}

#[test]
fn test_vb_buffer_byte_length() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(9) As Integer
        Console.WriteLine(Buffer.ByteLength(arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["40"]);
}

#[test]
fn test_vb_buffer_get_set_byte() {
    let src = r#"
Module Program
    Sub Main()
        Dim arr(1) As Integer
        Buffer.SetByte(arr, 0, 255)
        Dim b As Byte = Buffer.GetByte(arr, 0)
        Console.WriteLine(b)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_array_segment_to_array() {
    let src = r#"
Module Program
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        Dim segment As New ArraySegment(Of Integer)(numbers, 2, 2)
        Dim arr As Integer() = segment.ToArray()
        Console.WriteLine(arr.Length)
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "3,4"]);
}
