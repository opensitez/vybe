//! encoding/gob runtime: Register, Encode/Decode primitives, struct field skipping,
//! slice/map roundtrips, GobEncoder/GobDecoder interfaces — distinct from compile smokes
//! in `test_cover_encoding_extra.rs` and `test_stdlib_encoding_misc.rs`.

use crate::helpers::*;

go_run_cases! {
    gob_encode_decode_int_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; enc := gob.NewEncoder(&buf); enc.Encode(42); var v int; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["42"]
    ),
    gob_encode_decode_bool_true => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(true); var v bool; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["true"]
    ),
    gob_encode_decode_bool_false => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(false); var v bool; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["false"]
    ),
    gob_encode_decode_string_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(\"hello\"); var s string; gob.NewDecoder(&buf).Decode(&s); fmt.Println(s) }",
        vec!["hello"]
    ),
    gob_encode_decode_float64_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(3.5); var f float64; gob.NewDecoder(&buf).Decode(&f); fmt.Println(f) }",
        vec!["3.5"]
    ),
    gob_encode_decode_int_slice => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := []int{1, 2, 3}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back []int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(len(back)); fmt.Println(back[2]) }",
        vec!["3", "3"]
    ),
    gob_encode_decode_string_slice => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := []string{\"a\", \"b\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back []string; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[0]); fmt.Println(back[1]) }",
        vec!["a", "b"]
    ),
    gob_encode_decode_int_map => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := map[string]int{\"x\": 7, \"y\": 8}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[string]int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[\"x\"]); fmt.Println(back[\"y\"]) }",
        vec!["7", "8"]
    ),
    gob_encode_decode_string_map => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := map[int]string{1: \"one\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[int]string; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[1]) }",
        vec!["one"]
    ),
    gob_struct_exported_fields_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Pair struct { A int; B string }; func main() { orig := Pair{A: 10, B: \"go\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Pair; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.A); fmt.Println(back.B) }",
        vec!["10", "go"]
    ),
    gob_struct_unexported_field_skipped_on_encode => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Hidden struct { Pub int; priv int }; func main() { h := Hidden{Pub: 5, priv: 99}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(h); var back Hidden; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.Pub); fmt.Println(back.priv) }",
        vec!["5", "0"]
    ),
    gob_nested_struct_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Inner struct { V int }; type Outer struct { Inner Inner; Tag string }; func main() { orig := Outer{Inner: Inner{V: 4}, Tag: \"t\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Outer; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.Inner.V); fmt.Println(back.Tag) }",
        vec!["4", "t"]
    ),
    gob_pointer_to_int_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { n := 17; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(&n); var back *int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(*back) }",
        vec!["17"]
    ),
    gob_nil_pointer_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var p *int; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(p); var back *int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back == nil) }",
        vec!["true"]
    ),
    gob_encode_decode_uint8 => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(uint8(255)); var v uint8; gob.NewDecoder(&buf).Decode(&v); fmt.Println(int(v)) }",
        vec!["255"]
    ),
    gob_encode_decode_int64 => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(int64(1000000)); var v int64; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["1000000"]
    ),
    gob_empty_slice_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := []int{}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back []int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(len(back)) }",
        vec!["0"]
    ),
    gob_empty_map_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := map[string]int{}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[string]int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(len(back)) }",
        vec!["0"]
    ),
    gob_register_then_encode_custom_type => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type ID struct { N int }; func main() { gob.Register(ID{}); orig := ID{N: 3}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back ID; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.N) }",
        vec!["3"]
    ),
    gob_register_name_then_decode => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Tag struct { Label string }; func main() { gob.RegisterName(\"TagType\", Tag{}); orig := Tag{Label: \"x\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Tag; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.Label) }",
        vec!["x"]
    ),
    gob_encode_produces_non_empty_buffer => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(1); fmt.Println(buf.Len() > 0) }",
        vec!["true"]
    ),
    gob_two_values_sequential_decode => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; enc := gob.NewEncoder(&buf); enc.Encode(1); enc.Encode(2); dec := gob.NewDecoder(&buf); var a, b int; dec.Decode(&a); dec.Decode(&b); fmt.Println(a); fmt.Println(b) }",
        vec!["1", "2"]
    ),
    gob_struct_with_bool_field => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Flags struct { Ok bool; Count int }; func main() { orig := Flags{Ok: true, Count: 2}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Flags; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.Ok); fmt.Println(back.Count) }",
        vec!["true", "2"]
    ),
    gob_array_fixed_size_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := [3]int{4, 5, 6}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back [3]int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[0]); fmt.Println(back[2]) }",
        vec!["4", "6"]
    ),
    gob_map_with_multiple_entries => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := map[int]int{1: 10, 2: 20, 3: 30}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[int]int; gob.NewDecoder(&buf).Decode(&back); fmt.Println(len(back)); fmt.Println(back[2]) }",
        vec!["3", "20"]
    ),
    gob_interface_boxing_int => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(interface{}(99)); var v interface{}; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v.(int)) }",
        vec!["99"]
    ),
    gob_negative_int_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(-123); var v int; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["-123"]
    ),
    gob_zero_int_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(0); var v int; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["0"]
    ),
    gob_empty_string_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(\"\"); var s string; gob.NewDecoder(&buf).Decode(&s); fmt.Println(len(s)) }",
        vec!["0"]
    ),
    gob_struct_zero_values => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Empty struct { N int; S string }; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(Empty{}); var back Empty; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.N); fmt.Println(back.S) }",
        vec!["0", ""]
    ),
    gob_slice_of_structs => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Node struct { ID int }; func main() { orig := []Node{{ID: 1}, {ID: 2}}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back []Node; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[1].ID) }",
        vec!["2"]
    ),
    gob_map_string_to_struct => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Rec struct { Val int }; func main() { orig := map[string]Rec{\"k\": {Val: 9}}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[string]Rec; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[\"k\"].Val) }",
        vec!["9"]
    ),
    gob_gob_encoder_interface_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Custom struct { N int }; func (c Custom) GobEncode() ([]byte, error) { return []byte{byte(c.N)}, nil }; func (c *Custom) GobDecode(b []byte) error { c.N = int(b[0]); return nil }; func main() { orig := Custom{N: 7}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Custom; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.N) }",
        vec!["7"]
    ),
    gob_gob_decoder_mutates_receiver => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Box struct { Data []byte }; func (b *Box) GobDecode(p []byte) error { b.Data = append([]byte(nil), p...); return nil }; func (b Box) GobEncode() ([]byte, error) { return b.Data, nil }; func main() { orig := Box{Data: []byte(\"ab\")}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Box; gob.NewDecoder(&buf).Decode(&back); fmt.Println(string(back.Data)) }",
        vec!["ab"]
    ),
    gob_register_interface_value => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Counter struct { N int }; func main() { gob.Register(&Counter{}); orig := &Counter{N: 12}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back *Counter; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.N) }",
        vec!["12"]
    ),
    gob_float32_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(float32(1.25)); var v float32; gob.NewDecoder(&buf).Decode(&v); fmt.Println(v) }",
        vec!["1.25"]
    ),
    gob_byte_slice_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := []byte{1, 2, 3}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back []byte; gob.NewDecoder(&buf).Decode(&back); fmt.Println(len(back)); fmt.Println(int(back[1])) }",
        vec!["3", "2"]
    ),
    gob_struct_three_fields => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; type Triple struct { A int; B int; C string }; func main() { orig := Triple{A: 1, B: 2, C: \"c\"}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back Triple; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back.B); fmt.Println(back.C) }",
        vec!["2", "c"]
    ),
    gob_map_bool_keys_not_supported_use_int => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { orig := map[int]bool{0: false, 1: true}; var buf bytes.Buffer; gob.NewEncoder(&buf).Encode(orig); var back map[int]bool; gob.NewDecoder(&buf).Decode(&back); fmt.Println(back[1]) }",
        vec!["true"]
    ),
    gob_reencode_same_type_new_buffer => (
        "package main; import \"fmt\"; import \"encoding/gob\"; import \"bytes\"; func main() { var buf1, buf2 bytes.Buffer; gob.NewEncoder(&buf1).Encode(88); var v int; gob.NewDecoder(&buf1).Decode(&v); gob.NewEncoder(&buf2).Encode(v); var v2 int; gob.NewDecoder(&buf2).Decode(&v2); fmt.Println(v2) }",
        vec!["88"]
    ),
}

go_compile_cases! {
    gob_encoder_encode_value_reflect => "package main; import \"encoding/gob\"; import \"bytes\"; import \"reflect\"; func main() { e := gob.NewEncoder(bytes.NewBuffer(nil)); _ = e.EncodeValue(reflect.ValueOf(1)) }",
    gob_decoder_decode_value_reflect => "package main; import \"encoding/gob\"; import \"bytes\"; import \"reflect\"; func main() { d := gob.NewDecoder(bytes.NewBuffer(nil)); _ = d.DecodeValue(reflect.ValueOf(new(int)).Elem()) }",
    gob_gob_encoder_interface_compile => "package main; import \"encoding/gob\"; type T struct { N int }; func (t T) GobEncode() ([]byte, error) { return nil, nil }; func main() { var _ gob.GobEncoder = T{} }",
    gob_gob_decoder_interface_compile => "package main; import \"encoding/gob\"; type T struct { N int }; func (t *T) GobDecode([]byte) error { return nil }; func main() { var _ gob.GobDecoder = &T{} }",
    gob_register_multiple_types => "package main; import \"encoding/gob\"; func main() { gob.Register(int(0)); gob.Register(string(\"\")); gob.Register([]int{}) }",
    gob_register_name_custom => "package main; import \"encoding/gob\"; type Widget struct { X int }; func main() { gob.RegisterName(\"my.Widget\", Widget{}) }",
    gob_encode_struct_with_unexported => "package main; import \"encoding/gob\"; import \"bytes\"; type S struct { pub int; priv string }; func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(S{pub: 1, priv: \"x\"}) }",
    gob_decode_into_new_map => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { var m map[string]int; _ = gob.NewDecoder(bytes.NewBuffer(nil)).Decode(&m) }",
    gob_decode_into_new_slice => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { var s []string; _ = gob.NewDecoder(bytes.NewBuffer(nil)).Decode(&s) }",
    gob_nested_pointer_struct => "package main; import \"encoding/gob\"; import \"bytes\"; type Node struct { Next *Node; Val int }; func main() { n := &Node{Val: 1}; _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(n) }",
    gob_interface_slice_compile => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode([]interface{}{1, \"a\"}) }",
    gob_struct_embedded_anonymous => "package main; import \"encoding/gob\"; import \"bytes\"; type Base struct { ID int }; type Derived struct { Base; Name string }; func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(Derived{Base: Base{ID: 1}, Name: \"d\"}) }",
    gob_channel_not_encodable_compile => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { ch := make(chan int); _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(ch) }",
    gob_func_not_encodable_compile => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { f := func() {}; _ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(f) }",
    gob_complex128_roundtrip_compile => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { var buf bytes.Buffer; _ = gob.NewEncoder(&buf).Encode(complex(1, 2)); var c complex128; _ = gob.NewDecoder(&buf).Decode(&c) }",
}
