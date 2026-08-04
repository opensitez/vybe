//! `reflect` Value runtime semantics: TypeOf, ValueOf, Kind, Field, FieldByName,
//! SetInt, SetString, Call, CallSlice, Interface, IsNil, Elem, NumField, Method —
//! distinct from `test_reflect_unsafe_compile.rs` (compile smoke) and
//! `test_struct_tags_reflect.rs` (struct tag introspection).

go_run_cases! {
    reflect_typeof_int_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.TypeOf(42).Kind()) }",
        vec!["int"]
    ),
    reflect_typeof_string_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.TypeOf(\"go\").Kind()) }",
        vec!["string"]
    ),
    reflect_typeof_slice_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.TypeOf([]int{1}).Kind()) }",
        vec!["slice"]
    ),
    reflect_typeof_map_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.TypeOf(map[string]int{}).Kind()) }",
        vec!["map"]
    ),
    reflect_typeof_pointer_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { var x int; fmt.Println(reflect.TypeOf(&x).Kind()) }",
        vec!["ptr"]
    ),
    reflect_valueof_int_interface => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(7); fmt.Println(v.Int()) }",
        vec!["7"]
    ),
    reflect_valueof_string_interface => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(\"vybe\"); fmt.Println(v.String()) }",
        vec!["vybe"]
    ),
    reflect_valueof_bool_true => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(true); fmt.Println(v.Bool()) }",
        vec!["true"]
    ),
    reflect_kind_on_value => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.ValueOf(3.14).Kind()) }",
        vec!["float64"]
    ),
    reflect_numfield_struct => (
        "package main; import \"fmt\"; import \"reflect\"; type Person struct { Name string; Age int }; func main() { fmt.Println(reflect.TypeOf(Person{}).NumField()) }",
        vec!["2"]
    ),
    reflect_field_by_index_name => (
        "package main; import \"fmt\"; import \"reflect\"; type Pair struct { A int; B string }; func main() { f := reflect.TypeOf(Pair{}).Field(0); fmt.Println(f.Name) }",
        vec!["A"]
    ),
    reflect_field_by_index_type_kind => (
        "package main; import \"fmt\"; import \"reflect\"; type Pair struct { A int; B string }; func main() { f := reflect.TypeOf(Pair{}).Field(1); fmt.Println(f.Type.Kind()) }",
        vec!["string"]
    ),
    reflect_field_by_name_found => (
        "package main; import \"fmt\"; import \"reflect\"; type S struct { Score int }; func main() { f, ok := reflect.TypeOf(S{}).FieldByName(\"Score\"); fmt.Println(ok); fmt.Println(f.Name) }",
        vec!["true", "Score"]
    ),
    reflect_field_by_name_missing => (
        "package main; import \"fmt\"; import \"reflect\"; type S struct { X int }; func main() { _, ok := reflect.TypeOf(S{}).FieldByName(\"Missing\"); fmt.Println(ok) }",
        vec!["false"]
    ),
    reflect_setint_on_pointer_elem => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { x := 0; v := reflect.ValueOf(&x).Elem(); v.SetInt(42); fmt.Println(x) }",
        vec!["42"]
    ),
    reflect_setstring_on_pointer_elem => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { s := \"\"; v := reflect.ValueOf(&s).Elem(); v.SetString(\"hello\"); fmt.Println(s) }",
        vec!["hello"]
    ),
    reflect_setbool_on_pointer_elem => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { b := false; v := reflect.ValueOf(&b).Elem(); v.SetBool(true); fmt.Println(b) }",
        vec!["true"]
    ),
    reflect_elem_on_pointer_type => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { var x int; fmt.Println(reflect.TypeOf(&x).Elem().Kind()) }",
        vec!["int"]
    ),
    reflect_elem_on_value_pointer => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { x := 10; v := reflect.ValueOf(&x); fmt.Println(v.Elem().Int()) }",
        vec!["10"]
    ),
    reflect_isnil_on_nil_interface => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { var i interface{}; fmt.Println(reflect.ValueOf(i).IsNil()) }",
        vec!["true"]
    ),
    reflect_isnil_on_nil_slice_value => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { var s []int; fmt.Println(reflect.ValueOf(s).IsNil()) }",
        vec!["true"]
    ),
    reflect_isnil_on_int_not_valid => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.ValueOf(5).IsValid()) }",
        vec!["true"]
    ),
    reflect_interface_roundtrip_int => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(99); i := v.Interface().(int); fmt.Println(i) }",
        vec!["99"]
    ),
    reflect_interface_roundtrip_string => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(\"go\"); s := v.Interface().(string); fmt.Println(s) }",
        vec!["go"]
    ),
    reflect_call_no_arg_method => (
        "package main; import \"fmt\"; import \"reflect\"; type Counter struct { n int }; func (c *Counter) Inc() { c.n++ }; func (c Counter) Get() int { return c.n }; func main() { c := Counter{}; mv := reflect.ValueOf(&c).MethodByName(\"Get\"); out := mv.Call(nil); fmt.Println(out[0].Int()) }",
        vec!["0"]
    ),
    reflect_call_with_args => (
        "package main; import \"fmt\"; import \"reflect\"; func Add(a, b int) int { return a + b }; func main() { fv := reflect.ValueOf(Add); out := fv.Call([]reflect.Value{reflect.ValueOf(3), reflect.ValueOf(4)}); fmt.Println(out[0].Int()) }",
        vec!["7"]
    ),
    reflect_callslice_variadic => (
        "package main; import \"fmt\"; import \"reflect\"; func Sum(nums ...int) int { s := 0; for _, n := range nums { s += n }; return s }; func main() { fv := reflect.ValueOf(Sum); out := fv.CallSlice([]reflect.Value{reflect.ValueOf(1), reflect.ValueOf(2), reflect.ValueOf(3)}); fmt.Println(out[0].Int()) }",
        vec!["6"]
    ),
    reflect_method_by_name_inc => (
        "package main; import \"fmt\"; import \"reflect\"; type Box struct { V int }; func (b *Box) Set(v int) { b.V = v }; func main() { box := &Box{}; m := reflect.ValueOf(box).MethodByName(\"Set\"); m.Call([]reflect.Value{reflect.ValueOf(15)}); fmt.Println(box.V) }",
        vec!["15"]
    ),
    reflect_method_num_method => (
        "package main; import \"fmt\"; import \"reflect\"; type T struct{}; func (T) A() {}; func (T) B() {}; func (*T) C() {}; func main() { fmt.Println(reflect.TypeOf(T{}).NumMethod()); fmt.Println(reflect.TypeOf(&T{}).NumMethod()) }",
        vec!["2", "3"]
    ),
    reflect_value_len_slice => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf([]int{1, 2, 3}); fmt.Println(v.Len()) }",
        vec!["3"]
    ),
    reflect_value_index_slice => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf([]int{10, 20, 30}); fmt.Println(v.Index(1).Int()) }",
        vec!["20"]
    ),
    reflect_value_map_index => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { m := map[string]int{\"a\": 5}; v := reflect.ValueOf(m); fmt.Println(v.MapIndex(reflect.ValueOf(\"a\")).Int()) }",
        vec!["5"]
    ),
    reflect_value_can_set_on_elem => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { x := 1; v := reflect.ValueOf(&x).Elem(); fmt.Println(v.CanSet()) }",
        vec!["true"]
    ),
    reflect_value_can_set_on_copy => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(1); fmt.Println(v.CanSet()) }",
        vec!["false"]
    ),
    reflect_type_name_struct => (
        "package main; import \"fmt\"; import \"reflect\"; type Widget struct{}; func main() { fmt.Println(reflect.TypeOf(Widget{}).Name()) }",
        vec!["Widget"]
    ),
    reflect_ptr_to_struct_field_access => (
        "package main; import \"fmt\"; import \"reflect\"; type Data struct { N int }; func main() { d := &Data{N: 8}; v := reflect.ValueOf(d).Elem().Field(0); fmt.Println(v.Int()) }",
        vec!["8"]
    ),
    reflect_set_int_field_via_pointer => (
        "package main; import \"fmt\"; import \"reflect\"; type Data struct { N int }; func main() { d := &Data{}; f := reflect.ValueOf(d).Elem().Field(0); f.SetInt(33); fmt.Println(d.N) }",
        vec!["33"]
    ),
    reflect_call_method_changes_state => (
        "package main; import \"fmt\"; import \"reflect\"; type Acc struct { Sum int }; func (a *Acc) Add(n int) { a.Sum += n }; func main() { acc := &Acc{}; reflect.ValueOf(acc).MethodByName(\"Add\").Call([]reflect.Value{reflect.ValueOf(5)}); fmt.Println(acc.Sum) }",
        vec!["5"]
    ),
    reflect_value_float64 => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { v := reflect.ValueOf(2.5); fmt.Println(v.Float()) }",
        vec!["2.5"]
    ),
    reflect_typeof_func_kind => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.TypeOf(func() {}).Kind()) }",
        vec!["func"]
    ),
    reflect_value_is_zero_int => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.ValueOf(0).IsZero()) }",
        vec!["true"]
    ),
    reflect_value_is_zero_nonempty_string => (
        "package main; import \"fmt\"; import \"reflect\"; func main() { fmt.Println(reflect.ValueOf(\"x\").IsZero()) }",
        vec!["false"]
    ),
    reflect_field_by_name_func => (
        "package main; import \"fmt\"; import \"reflect\"; type Row struct { Alpha int; Beta int; Gamma int }; func main() { f, ok := reflect.TypeOf(Row{}).FieldByNameFunc(func(name string) bool { return len(name) == 5 }); fmt.Println(ok); fmt.Println(f.Name) }",
        vec!["true", "Alpha"]
    ),
}

go_compile_cases! {
    reflect_typeof_chan => "package main; import \"reflect\"; func main() { _ = reflect.TypeOf(make(chan int)).Kind() }",
    reflect_valueof_struct => "package main; import \"reflect\"; type S struct { X int }; func main() { _ = reflect.ValueOf(S{X: 1}).Field(0) }",
    reflect_set_uint_on_elem => "package main; import \"reflect\"; func main() { var x uint; v := reflect.ValueOf(&x).Elem(); v.SetUint(7) }",
    reflect_call_two_returns => "package main; import \"reflect\"; func DivMod(a, b int) (int, int) { return a / b, a % b }; func main() { out := reflect.ValueOf(DivMod).Call([]reflect.Value{reflect.ValueOf(10), reflect.ValueOf(3)}); _, _ = out[0], out[1] }",
    reflect_method_value_interface => "package main; import \"reflect\"; type R struct{}; func (R) M() string { return \"ok\" }; func main() { _ = reflect.ValueOf(R{}).Method(0).Interface() }",
    reflect_elem_on_slice_header => "package main; import \"reflect\"; func main() { s := []int{1}; _ = reflect.ValueOf(s).Index(0).Interface() }",
    reflect_pointer_interface => "package main; import \"reflect\"; func main() { x := 1; _ = reflect.ValueOf(&x).Interface().(*int) }",
    reflect_struct_field_anonymous => "package main; import \"reflect\"; type Inner struct { N int }; type Outer struct { Inner }; func main() { _ = reflect.TypeOf(Outer{}).Field(0) }",
    reflect_array_type_num_field_zero => "package main; import \"reflect\"; func main() { _ = reflect.TypeOf([3]int{}).NumField() }",
    reflect_map_set_via_reflect => "package main; import \"reflect\"; func main() { m := map[string]int{}; mv := reflect.ValueOf(m); mv.SetMapIndex(reflect.ValueOf(\"k\"), reflect.ValueOf(1)) }",
    reflect_slice_append_reflect => "package main; import \"reflect\"; func main() { s := []int{1}; sv := reflect.ValueOf(&s).Elem(); sv.Set(reflect.Append(sv, reflect.ValueOf(2))) }",
    reflect_func_type_in_num_in => "package main; import \"reflect\"; func main() { t := reflect.TypeOf(func(int, string) bool { return true }); _ = t.NumIn() }",
    reflect_func_type_num_out => "package main; import \"reflect\"; func main() { t := reflect.TypeOf(func() (int, error) { return 0, nil }); _ = t.NumOut() }",
    reflect_interface_implemented_check => "package main; import \"reflect\"; type I interface { F() }; type T struct{}; func (T) F() {}; func main() { _ = reflect.TypeOf(T{}).Implements(reflect.TypeOf((*I)(nil)).Elem()) }",
    reflect_value_convert_int_to_int64 => "package main; import \"reflect\"; func main() { v := reflect.ValueOf(int(5)); _ = v.Convert(reflect.TypeOf(int64(0))) }",
    reflect_new_allocates => "package main; import \"reflect\"; func main() { p := reflect.New(reflect.TypeOf(0)); _ = p.Elem().SetInt(1) }",
    reflect_indirect_on_pointer => "package main; import \"reflect\"; func main() { x := 3; _ = reflect.Indirect(reflect.ValueOf(&x)).Int() }",
    reflect_visible_field_count => "package main; import \"reflect\"; type S struct { Pub int; priv int }; func main() { _ = reflect.TypeOf(S{}).NumField() }",
    reflect_method_by_name_not_found => "package main; import \"reflect\"; type T struct{}; func main() { _ = reflect.ValueOf(T{}).MethodByName(\"Missing\").IsValid() }",
    reflect_call_invalid_panic_compile => "package main; import \"reflect\"; func noop() {}; func main() { _ = reflect.ValueOf(noop).Call(nil) }",
    reflect_struct_tag_on_field => "package main; import \"reflect\"; type T struct { X int `json:\"x\"` }; func main() { _ = reflect.TypeOf(T{}).Field(0).Tag.Get(\"json\") }",
    reflect_value_bytes => "package main; import \"reflect\"; func main() { _ = reflect.ValueOf([]byte{'a'}).Bytes() }",
    reflect_value_map_keys => "package main; import \"reflect\"; func main() { _ = reflect.ValueOf(map[int]int{1: 1}).MapKeys() }",
}
