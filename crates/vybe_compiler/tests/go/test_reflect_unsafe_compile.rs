//! reflect and unsafe: compile-time type introspection and size/align.

use crate::helpers::*;

go_compile_cases! {
    reflect_type_of_int => "package main; import \"reflect\"; func main() { _ = reflect.TypeOf(0) }",
    reflect_value_of_string => "package main; import \"reflect\"; func main() { _ = reflect.ValueOf(\"x\") }",
    reflect_kind_slice => "package main; import \"reflect\"; func main() { _ = reflect.TypeOf([]int{}).Kind() }",
    reflect_struct_field_count => "package main; import \"reflect\"; type S struct { A int; B string }; func main() { _ = reflect.TypeOf(S{}).NumField() }",
    reflect_ptr_elem => "package main; import \"reflect\"; func main() { var x int; _ = reflect.TypeOf(&x).Elem() }",
    reflect_interface_implemented => "package main; import \"reflect\"; type R interface { M() }; type T struct{}; func (T) M() {}; func main() { _ = reflect.TypeOf((*T)(nil)).Implements(reflect.TypeOf((*R)(nil)).Elem()) }",
    unsafe_sizeof_int => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(int(0)) }",
    unsafe_alignof_struct => "package main; import \"unsafe\"; type S struct { a int8; b int32 }; func main() { _ = unsafe.Alignof(S{}) }",
    unsafe_offsetof_field => "package main; import \"unsafe\"; type S struct { a int; b int }; func main() { _ = unsafe.Offsetof(S{}.b) }",
    uintptr_convert_pointer => "package main; import \"unsafe\"; func main() { var x int; p := &x; _ = uintptr(unsafe.Pointer(p)) }",
}
