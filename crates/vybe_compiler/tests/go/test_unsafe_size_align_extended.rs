//! unsafe: Sizeof, Alignof, Offsetof on structs, Pointer, SliceData, StringData,
//! uintptr conversions, and safe pointer patterns without arithmetic — extended
//! coverage distinct from `test_embed_unsafe_size.rs` and `test_reflect_unsafe_compile.rs`.

go_run_cases! {
    unsafe_sizeof_int8_one => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(int8(0))) }",
        vec!["1"]
    ),
    unsafe_sizeof_int32_four => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(int32(0))) }",
        vec!["4"]
    ),
    unsafe_sizeof_int64_eight => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(int64(0))) }",
        vec!["8"]
    ),
    unsafe_sizeof_bool => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(true)) }",
        vec!["1"]
    ),
    unsafe_sizeof_float64 => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(float64(0))) }",
        vec!["8"]
    ),
    unsafe_sizeof_string_header => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof(\"\")) }",
        vec!["16"]
    ),
    unsafe_sizeof_slice_header => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof([]int{})) }",
        vec!["24"]
    ),
    unsafe_sizeof_pointer => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { var p *int; fmt.Println(unsafe.Sizeof(p)) }",
        vec!["8"]
    ),

    unsafe_sizeof_struct_padded => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a byte; b int64 }; func main() { fmt.Println(unsafe.Sizeof(S{})) }",
        vec!["16"]
    ),
    unsafe_sizeof_struct_no_padding => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a int32; b int32 }; func main() { fmt.Println(unsafe.Sizeof(S{})) }",
        vec!["8"]
    ),
    unsafe_sizeof_struct_three_fields => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { x int16; y int16; z int32 }; func main() { fmt.Println(unsafe.Sizeof(S{})) }",
        vec!["8"]
    ),
    unsafe_sizeof_nested_struct => (
        "package main; import \"fmt\"; import \"unsafe\"; type Inner struct { v int32 }; type Outer struct { i Inner; flag bool }; func main() { fmt.Println(unsafe.Sizeof(Outer{})) }",
        vec!["8"]
    ),
    unsafe_sizeof_array_fixed => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Sizeof([4]int32{})) }",
        vec!["16"]
    ),
    unsafe_sizeof_empty_struct => (
        "package main; import \"fmt\"; import \"unsafe\"; type E struct{}; func main() { fmt.Println(unsafe.Sizeof(E{})) }",
        vec!["0"]
    ),

    unsafe_alignof_int64 => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Alignof(int64(0))) }",
        vec!["8"]
    ),
    unsafe_alignof_int32 => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Alignof(int32(0))) }",
        vec!["4"]
    ),
    unsafe_alignof_byte => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Alignof(byte(0))) }",
        vec!["1"]
    ),
    unsafe_alignof_struct_max_field => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a byte; b int64 }; func main() { fmt.Println(unsafe.Alignof(S{})) }",
        vec!["8"]
    ),
    unsafe_alignof_string => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { fmt.Println(unsafe.Alignof(\"\")) }",
        vec!["8"]
    ),

    unsafe_offsetof_first_field_zero => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a int; b int }; func main() { fmt.Println(unsafe.Offsetof(S{}.a)) }",
        vec!["0"]
    ),
    unsafe_offsetof_second_int_field => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a int; b int }; func main() { fmt.Println(unsafe.Offsetof(S{}.b)) }",
        vec!["8"]
    ),
    unsafe_offsetof_after_byte_padding => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { a byte; b int64 }; func main() { fmt.Println(unsafe.Offsetof(S{}.b)) }",
        vec!["8"]
    ),
    unsafe_offsetof_string_field => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { name string; n int }; func main() { fmt.Println(unsafe.Offsetof(S{}.n)) }",
        vec!["16"]
    ),
    unsafe_offsetof_slice_field => (
        "package main; import \"fmt\"; import \"unsafe\"; type S struct { data []int; tag byte }; func main() { fmt.Println(unsafe.Offsetof(S{}.tag)) }",
        vec!["24"]
    ),

    unsafe_pointer_from_int_var => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { var x int = 7; p := unsafe.Pointer(&x); fmt.Println(p != nil) }",
        vec!["true"]
    ),
    unsafe_pointer_from_array_element => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { arr := [2]int{1, 2}; p := unsafe.Pointer(&arr[1]); fmt.Println(p != nil) }",
        vec!["true"]
    ),
    unsafe_pointer_nil_compare => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { var p unsafe.Pointer; fmt.Println(p == nil) }",
        vec!["true"]
    ),

    uintptr_from_pointer_roundtrip_nonzero => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { var x int = 3; u := uintptr(unsafe.Pointer(&x)); p := unsafe.Pointer(u); fmt.Println(p != nil) }",
        vec!["true"]
    ),
    uintptr_zero_from_nil_pointer => (
        "package main; import \"fmt\"; import \"unsafe\"; func main() { var p *int; fmt.Println(uintptr(unsafe.Pointer(p)) == 0) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    unsafe_sizeof_complex128 => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(complex128(0)) }",
    unsafe_sizeof_func_value => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(func(){}) }",
    unsafe_sizeof_interface_value => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(interface{}(nil)) }",
    unsafe_sizeof_map_header => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(map[string]int{}) }",
    unsafe_sizeof_chan_header => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(make(chan int)) }",
    unsafe_sizeof_struct_bool_int16 => "package main; import \"unsafe\"; type S struct { ok bool; n int16 }; func main() { _ = unsafe.Sizeof(S{}) }",
    unsafe_sizeof_struct_float32_padding => "package main; import \"unsafe\"; type S struct { f float32; b byte }; func main() { _ = unsafe.Sizeof(S{}) }",
    unsafe_sizeof_struct_embedded_anonymous => "package main; import \"unsafe\"; type Base struct { id int32 }; type Child struct { Base; name string }; func main() { _ = unsafe.Sizeof(Child{}) }",

    unsafe_alignof_complex64 => "package main; import \"unsafe\"; func main() { _ = unsafe.Alignof(complex64(0)) }",
    unsafe_alignof_float64 => "package main; import \"unsafe\"; func main() { _ = unsafe.Alignof(float64(0)) }",
    unsafe_alignof_pointer_type => "package main; import \"unsafe\"; func main() { var p *byte; _ = unsafe.Alignof(p) }",
    unsafe_alignof_array_type => "package main; import \"unsafe\"; func main() { _ = unsafe.Alignof([8]byte{}) }",
    unsafe_alignof_struct_mixed => "package main; import \"unsafe\"; type S struct { a int16; b int32; c byte }; func main() { _ = unsafe.Alignof(S{}) }",

    unsafe_offsetof_third_field => "package main; import \"unsafe\"; type S struct { a byte; b byte; c int32 }; func main() { _ = unsafe.Offsetof(S{}.c) }",
    unsafe_offsetof_bool_after_int => "package main; import \"unsafe\"; type S struct { n int64; flag bool }; func main() { _ = unsafe.Offsetof(S{}.flag) }",
    unsafe_offsetof_pointer_field => "package main; import \"unsafe\"; type S struct { p *int; x int }; func main() { _ = unsafe.Offsetof(S{}.x) }",
    unsafe_offsetof_array_field => "package main; import \"unsafe\"; type S struct { buf [4]byte; n int }; func main() { _ = unsafe.Offsetof(S{}.n) }",
    unsafe_offsetof_embedded_base_field => "package main; import \"unsafe\"; type Base struct { id int }; type Wrap struct { Base; extra byte }; func main() { _ = unsafe.Offsetof(Wrap{}.extra) }",

    unsafe_pointer_from_string_data => "package main; import \"unsafe\"; func main() { s := \"go\"; _ = unsafe.Pointer(unsafe.StringData(s)) }",
    unsafe_pointer_from_slice_data => "package main; import \"unsafe\"; func main() { sl := []int{1,2}; _ = unsafe.Pointer(unsafe.SliceData(sl)) }",
    unsafe_pointer_from_struct_field => "package main; import \"unsafe\"; type S struct { n int }; func main() { var s S; _ = unsafe.Pointer(&s.n) }",
    unsafe_pointer_convert_to_byte_ptr => "package main; import \"unsafe\"; func main() { var x int; _ = (*byte)(unsafe.Pointer(&x)) }",

    unsafe_slice_data_nonempty => "package main; import \"unsafe\"; func main() { sl := []byte(\"abc\"); _ = unsafe.SliceData(sl) }",
    unsafe_slice_data_nil => "package main; import \"unsafe\"; func main() { var sl []int; _ = unsafe.SliceData(sl) }",
    unsafe_string_data_nonempty => "package main; import \"unsafe\"; func main() { _ = unsafe.StringData(\"vybe\") }",
    unsafe_string_data_empty => "package main; import \"unsafe\"; func main() { _ = unsafe.StringData(\"\") }",
    unsafe_string_from_bytes_zero_len => "package main; import \"unsafe\"; func main() { _ = unsafe.String((*byte)(nil), 0) }",
    unsafe_string_from_slice_data => "package main; import \"unsafe\"; func main() { b := []byte(\"x\"); _ = unsafe.String(unsafe.SliceData(b), len(b)) }",
    unsafe_slice_from_pointer_zero_len => "package main; import \"unsafe\"; func main() { _ = unsafe.Slice((*int)(nil), 0) }",

    uintptr_from_string_data => "package main; import \"unsafe\"; func main() { s := \"a\"; _ = uintptr(unsafe.Pointer(unsafe.StringData(s))) }",
    uintptr_from_slice_data => "package main; import \"unsafe\"; func main() { sl := []byte{1}; _ = uintptr(unsafe.Pointer(unsafe.SliceData(sl))) }",
    uintptr_to_pointer_int => "package main; import \"unsafe\"; func main() { var x int; u := uintptr(unsafe.Pointer(&x)); _ = (*int)(unsafe.Pointer(u)) }",
    uintptr_to_pointer_byte => "package main; import \"unsafe\"; func main() { var b byte; u := uintptr(unsafe.Pointer(&b)); _ = (*byte)(unsafe.Pointer(u)) }",

    no_arithmetic_pointer_to_struct => "package main; import \"unsafe\"; type Node struct { next *Node }; func main() { var n Node; _ = unsafe.Pointer(&n) }",
    no_arithmetic_pointer_store_load => "package main; import \"unsafe\"; func main() { var x int; p := (*int)(unsafe.Pointer(&x)); *p = 1; _ = *p }",
    no_arithmetic_slice_data_read => "package main; import \"unsafe\"; func main() { b := []byte{10}; ptr := unsafe.SliceData(b); _ = *ptr }",
    no_arithmetic_string_data_len => "package main; import \"unsafe\"; func main() { s := \"go\"; ptr := unsafe.StringData(s); _ = ptr; _ = len(s) }",
    no_arithmetic_interface_boxed => "package main; import \"unsafe\"; func main() { var v interface{} = 42; _ = unsafe.Pointer(&v) }",
    no_arithmetic_func_pointer => "package main; import \"unsafe\"; func main() { f := func() {}; _ = unsafe.Pointer(&f) }",
    no_arithmetic_array_element_addr => "package main; import \"unsafe\"; func main() { a := [3]int{1,2,3}; _ = unsafe.Pointer(&a[2]) }",
    no_arithmetic_nested_struct_addr => "package main; import \"unsafe\"; type S struct { inner [2]byte }; func main() { var s S; _ = unsafe.Pointer(&s.inner[1]) }",
    no_arithmetic_chan_addr => "package main; import \"unsafe\"; func main() { ch := make(chan int, 1); _ = unsafe.Pointer(&ch) }",
    no_arithmetic_map_var_addr => "package main; import \"unsafe\"; func main() { m := map[string]int{\"a\":1}; _ = unsafe.Pointer(&m) }",
    no_arithmetic_uintptr_compare_zero => "package main; import \"unsafe\"; func main() { var p *int; _ = uintptr(unsafe.Pointer(p)) == 0 }",
    no_arithmetic_double_pointer_convert => "package main; import \"unsafe\"; func main() { var x int; p := &x; _ = unsafe.Pointer(p) }",
}
