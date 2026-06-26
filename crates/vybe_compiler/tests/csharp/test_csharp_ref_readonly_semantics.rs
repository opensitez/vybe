//! `ref readonly` returns, `readonly ref struct` rules, and `Memory<T>` via print.
//! GAP: memory/ref-readonly coverage is thin in the existing suite.

use crate::csharp_cases;

csharp_cases! {
    ref_readonly_return_reads_array_element_without_copy => {
        r#"int[] data={10,20,30}; ref readonly int Peek(int i)=>ref data[i]; Console.WriteLine(Peek(1));"#,
        ["20"]
    };

    ref_readonly_return_from_local_function => {
        r#"int total=5; ref readonly int View(){return ref total;} Console.WriteLine(View());"#,
        ["5"]
    };

    ref_readonly_local_aliases_value_without_writable_ref => {
        r#"int[] nums={1,2,3}; ref readonly int slot=ref nums[2]; Console.WriteLine(slot);"#,
        ["3"]
    };

    ref_readonly_parameter_accepts_array_element_reference => {
        r#"void Show(ref readonly int value){Console.WriteLine(value);} int[] arr={7,8}; Show(ref arr[0]);"#,
        ["7"]
    };

    ref_readonly_field_in_readonly_struct => {
        r#"readonly struct Pair{public readonly int First; public readonly int Second; public Pair(int a,int b){First=a; Second=b;}} var p=new Pair(2,3); Console.WriteLine(p.First+p.Second);"#,
        ["5"]
    };

    readonly_ref_struct_can_be_local_variable => {
        r#"readonly ref struct Marker{public readonly int Code; public Marker(int c){Code=c;}} var m=new Marker(42); Console.WriteLine(m.Code);"#,
        ["42"]
    };

    readonly_ref_struct_method_reads_field => {
        r#"readonly ref struct Counter{public readonly int Value; public Counter(int v){Value=v;} public int Doubled()=>Value*2;} var c=new Counter(6); Console.WriteLine(c.Doubled());"#,
        ["12"]
    };

    ref_struct_is_value_type_not_reference => {
        r#"ref struct S{public int N;} var a=new S(); a.N=1; var b=a; b.N=2; Console.WriteLine(a.N);"#,
        ["1"]
    };

    ref_struct_copy_is_independent_for_value_fields => {
        r#"ref struct Box{public int Item;} var x=new Box(); x.Item=10; var y=x; y.Item=99; Console.WriteLine(x.Item);"#,
        ["10"]
    };

    ref_readonly_return_chained_to_readonly_local => {
        r#"int[] values={100,200}; ref readonly int Head()=>ref values[0]; ref readonly int h=ref Head(); Console.WriteLine(h);"#,
        ["100"]
    };

    ref_readonly_indexer_return_on_array_wrapper => {
        r#"class Buffer{private int[] _data={5,6,7}; public ref readonly int this[int i]=>ref _data[i];} var b=new Buffer(); Console.WriteLine(b[2]);"#,
        ["7"]
    };

    ref_readonly_property_returns_field_by_readonly_ref => {
        r#"struct Point{public int X; public ref readonly int Rx=>ref X;} var p=new Point(); p.X=11; Console.WriteLine(p.Rx);"#,
        ["11"]
    };

    ref_readonly_cannot_be_reassigned_but_source_mutation_visible => {
        r#"int[] arr={1,2}; ref readonly int r=ref arr[1]; arr[1]=50; Console.WriteLine(r);"#,
        ["50"]
    };

    memory_constructor_from_array_prints_span_length => {
        r#"var memory=new System.Memory<int>(new int[]{1,2,3}); Console.WriteLine(memory.Length);"#,
        ["3"]
    };

    memory_span_prints_first_element => {
        r#"var memory=new System.Memory<int>(new int[]{4,5,6}); Console.WriteLine(memory.Span[0]);"#,
        ["4"]
    };

    memory_span_prints_last_element => {
        r#"var memory=new System.Memory<int>(new int[]{4,5,6}); Console.WriteLine(memory.Span[2]);"#,
        ["6"]
    };

    memory_span_slice_prints_middle_element => {
        r#"var memory=new System.Memory<int>(new int[]{10,20,30,40}); Console.WriteLine(memory.Span.Slice(1,2)[1]);"#,
        ["30"]
    };

    memory_span_write_updates_printed_value => {
        r#"var memory=new System.Memory<int>(new int[]{1,2,3}); memory.Span[1]=88; Console.WriteLine(memory.Span[1]);"#,
        ["88"]
    };

    memory_is_empty_false_for_non_empty_buffer => {
        r#"var memory=new System.Memory<int>(new int[]{1}); Console.WriteLine(memory.IsEmpty);"#,
        ["False"]
    };

    memory_is_empty_true_for_empty_buffer => {
        r#"var memory=new System.Memory<int>(new int[]{}); Console.WriteLine(memory.IsEmpty);"#,
        ["True"]
    };

    memory_slice_offset_prints_element_at_new_origin => {
        r#"var memory=new System.Memory<int>(new int[]{2,4,6,8}); Console.WriteLine(memory.Slice(2).Span[0]);"#,
        ["6"]
    };

    memory_slice_offset_and_length_prints_inner_value => {
        r#"var memory=new System.Memory<int>(new int[]{2,4,6,8}); Console.WriteLine(memory.Slice(1,2).Span[1]);"#,
        ["6"]
    };

    readonly_memory_span_from_string_chars_prints_length => {
        r#"System.ReadOnlyMemory<char> mem="hello".AsMemory(); Console.WriteLine(mem.Length);"#,
        ["5"]
    };

    readonly_memory_span_from_string_prints_first_char => {
        r#"System.ReadOnlyMemory<char> mem="hello".AsMemory(); Console.WriteLine(mem.Span[0]);"#,
        ["104"]
    };

    memory_to_array_prints_copied_element => {
        r#"var memory=new System.Memory<int>(new int[]{9,8,7}); Console.WriteLine(memory.ToArray()[2]);"#,
        ["7"]
    };

    ref_readonly_in_foreach_reads_current_element => {
        r#"int[] data={3,6,9}; int sum=0; foreach(ref readonly int n in data){sum+=n;} Console.WriteLine(sum);"#,
        ["18"]
    };

    ref_readonly_return_of_struct_field => {
        r#"struct Widget{public int Id;} Widget w=new Widget(); w.Id=7; ref readonly int Get(ref Widget item)=>ref item.Id; Console.WriteLine(Get(ref w));"#,
        ["7"]
    };

    readonly_ref_struct_static_readonly_field => {
        r#"readonly ref struct Limits{public static readonly int Max=128; public int Value;} Console.WriteLine(Limits.Max);"#,
        ["128"]
    };

    ref_readonly_out_parameter_pattern_via_method => {
        r#"bool TryGet(ref readonly int[] src,int i,out int value){value=src[i]; return true;} int[] arr={12}; TryGet(ref arr,0,out int v); Console.WriteLine(v);"#,
        ["12"]
    };

    ref_readonly_extension_method_on_span => {
        r#"static class SpanExt{public static int First(ref readonly System.Span<int> span)=>span.Length>0?span[0]:-1;} System.Span<int> s=stackalloc int[2]{5,6}; Console.WriteLine(SpanExt.First(ref s));"#,
        ["5"]
    };

    memory_span_sequence_equal_prints_true_for_same_values => {
        r#"var left=new System.Memory<int>(new int[]{1,2}); var right=new System.Memory<int>(new int[]{1,2}); Console.WriteLine(left.Span.SequenceEqual(right.Span));"#,
        ["True"]
    };

    memory_span_sequence_equal_prints_false_for_different_values => {
        r#"var left=new System.Memory<int>(new int[]{1,2}); var right=new System.Memory<int>(new int[]{1,9}); Console.WriteLine(left.Span.SequenceEqual(right.Span));"#,
        ["False"]
    };

    ref_readonly_conditional_selects_branch_reference => {
        r#"int[] arr={1,2,3}; bool pickSecond=true; ref readonly int chosen=ref (pickSecond?ref arr[1]:ref arr[0]); Console.WriteLine(chosen);"#,
        ["2"]
    };

    readonly_ref_struct_equality_by_value => {
        r#"readonly ref struct Tag{public readonly int Id; public Tag(int id){Id=id;} public bool Equals(Tag other)=>Id==other.Id;} var a=new Tag(1); var b=new Tag(1); Console.WriteLine(a.Equals(b));"#,
        ["True"]
    };

    readonly_ref_struct_inequality_when_ids_differ => {
        r#"readonly ref struct Tag{public readonly int Id; public Tag(int id){Id=id;} public bool Equals(Tag other)=>Id==other.Id;} var a=new Tag(1); var b=new Tag(2); Console.WriteLine(a.Equals(b));"#,
        ["False"]
    };

    memory_span_index_from_end_prints_last => {
        r#"var memory=new System.Memory<int>(new int[]{10,20,30}); Console.WriteLine(memory.Span[^1]);"#,
        ["30"]
    };

    memory_span_clear_zeroes_elements => {
        r#"var memory=new System.Memory<int>(new int[]{5,6}); memory.Span.Clear(); Console.WriteLine(memory.Span[0]);"#,
        ["0"]
    };

    memory_span_fill_sets_uniform_value => {
        r#"var memory=new System.Memory<int>(new int[]{0,0,0}); memory.Span.Fill(4); Console.WriteLine(memory.Span[2]);"#,
        ["4"]
    };

    ref_readonly_nested_property_chain => {
        r#"class Node{public int Value;} class Holder{public Node Inner=new Node();} var h=new Holder(); h.Inner.Value=33; ref readonly int Read(ref Holder host)=>ref host.Inner.Value; Console.WriteLine(Read(ref h));"#,
        ["33"]
    };

    memory_length_matches_array_after_slice => {
        r#"var memory=new System.Memory<int>(new int[]{1,2,3,4,5}); Console.WriteLine(memory.Slice(2).Length);"#,
        ["3"]
    };

    ref_readonly_in_operator_overload_context => {
        r#"readonly struct Num{public readonly int Value; public Num(int v){Value=v;} public static bool operator ==(Num a, ref readonly Num b)=>a.Value==b.Value; public static bool operator !=(Num a, ref readonly Num b)=>!(a==b);} var x=new Num(4); var y=new Num(4); Console.WriteLine(x==ref y);"#,
        ["True"]
    };

    memory_span_copy_to_prints_destination_value => {
        r#"var src=new System.Memory<int>(new int[]{7,8}); int[] dst=new int[2]; src.Span.CopyTo(dst); Console.WriteLine(dst[1]);"#,
        ["8"]
    };

    ref_readonly_from_static_readonly_field => {
        r#"class Defaults{public static readonly int Seed=15;} ref readonly int SeedRef()=>ref Defaults.Seed; Console.WriteLine(SeedRef());"#,
        ["15"]
    };

    readonly_ref_struct_with_multiple_readonly_fields => {
        r#"readonly ref struct Rect{public readonly int W; public readonly int H; public Rect(int w,int h){W=w; H=h;} public int Area()=>W*H;} var r=new Rect(3,4); Console.WriteLine(r.Area());"#,
        ["12"]
    };

    memory_span_try_copy_to_reports_success => {
        r#"var src=new System.Memory<int>(new int[]{1,2}); int[] dst=new int[2]; Console.WriteLine(src.Span.TryCopyTo(dst));"#,
        ["True"]
    };

    ref_readonly_parameter_in_generic_method => {
        r#"static int Read<T>(ref readonly T value) where T: struct { return value.ToString().Length; } int n=123; Console.WriteLine(Read(ref n)>0);"#,
        ["True"]
    };

    memory_span_contains_value_prints_true => {
        r#"var memory=new System.Memory<int>(new int[]{2,4,6}); Console.WriteLine(memory.Span.Contains(4));"#,
        ["True"]
    };

    ref_readonly_return_does_not_copy_large_struct_field => {
        r#"struct Big{public int A; public int B; public int C;} Big item=new Big(); item.B=77; ref readonly int Read(ref Big target)=>ref target.B; Console.WriteLine(Read(ref item));"#,
        ["77"]
    };
}
