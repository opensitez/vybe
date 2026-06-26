//! encoding/json: Marshal and Unmarshal — primitives, structs, tags, slices, maps, null, nested.

use crate::helpers::*;

go_run_cases! {
    marshal_bool_true => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(true); fmt.Println(string(b)) }", vec!["true"]),
    marshal_bool_false => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(false); fmt.Println(string(b)) }", vec!["false"]),
    marshal_int_zero => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(0); fmt.Println(string(b)) }", vec!["0"]),
    marshal_int_negative => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(-7); fmt.Println(string(b)) }", vec!["-7"]),
    marshal_float_decimal => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(1.5); fmt.Println(string(b)) }", vec!["1.5"]),
    marshal_string_quoted => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(\"hello\"); fmt.Println(string(b)) }", vec!["\"hello\""]),
    marshal_nil_interface => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(nil); fmt.Println(string(b)) }", vec!["null"]),
    marshal_struct_two_fields => ("package main; import \"fmt\"; import \"encoding/json\"; type Person struct { Name string; Age int }; func main() { b, _ := json.Marshal(Person{Name: \"Bob\", Age: 30}); fmt.Println(string(b)) }", vec!["{\"Name\":\"Bob\",\"Age\":30}"]),
    marshal_struct_omitempty_skips_zero => ("package main; import \"fmt\"; import \"encoding/json\"; type Data struct { Count int `json:\",omitempty\"`; Label string `json:\",omitempty\"` }; func main() { b, _ := json.Marshal(Data{}); fmt.Println(string(b)) }", vec!["{}"]),
    marshal_struct_renames_with_tag => ("package main; import \"fmt\"; import \"encoding/json\"; type Item struct { ID int `json:\"id\"` }; func main() { b, _ := json.Marshal(Item{ID: 1}); fmt.Println(string(b)) }", vec!["{\"id\":1}"]),
    marshal_struct_dash_omits_field => ("package main; import \"fmt\"; import \"encoding/json\"; type Config struct { Secret string `json:\"-\"`; OK bool }; func main() { b, _ := json.Marshal(Config{Secret: \"hidden\", OK: true}); fmt.Println(string(b)) }", vec!["{\"OK\":true}"]),
    marshal_slice_ints => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal([]int{1, 2, 3}); fmt.Println(string(b)) }", vec!["[1,2,3]"]),
    marshal_nil_slice_null => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []int; b, _ := json.Marshal(s); fmt.Println(string(b)) }", vec!["null"]),
    marshal_empty_slice_brackets => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal([]int{}); fmt.Println(string(b)) }", vec!["[]"]),
    marshal_map_one_entry => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(map[string]int{\"a\": 1}); fmt.Println(string(b)) }", vec!["{\"a\":1}"]),
    marshal_nil_map_null => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var m map[string]int; b, _ := json.Marshal(m); fmt.Println(string(b)) }", vec!["null"]),
    marshal_nil_pointer_null => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var p *int; b, _ := json.Marshal(p); fmt.Println(string(b)) }", vec!["null"]),
    marshal_pointer_dereferences => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { n := 9; b, _ := json.Marshal(&n); fmt.Println(string(b)) }", vec!["9"]),
    marshal_nested_object => ("package main; import \"fmt\"; import \"encoding/json\"; type Child struct { N int }; type Parent struct { Child Child; Tag string }; func main() { b, _ := json.Marshal(Parent{Child: Child{N: 1}, Tag: \"x\"}); fmt.Println(string(b)) }", vec!["{\"Child\":{\"N\":1},\"Tag\":\"x\"}"]),
    unmarshal_bool_true => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var v bool; json.Unmarshal([]byte(\"true\"), &v); fmt.Println(v) }", vec!["true"]),
    unmarshal_int_from_number => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var n int; json.Unmarshal([]byte(\"42\"), &n); fmt.Println(n) }", vec!["42"]),
    unmarshal_string_from_quoted => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var s string; json.Unmarshal([]byte(\"\\\"hi\\\"\"), &s); fmt.Println(s) }", vec!["hi"]),
    unmarshal_null_nil_pointer => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var p *int; json.Unmarshal([]byte(\"null\"), &p); fmt.Println(p == nil) }", vec!["true"]),
    unmarshal_null_int_becomes_zero => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var n int; json.Unmarshal([]byte(\"null\"), &n); fmt.Println(n) }", vec!["0"]),
    unmarshal_struct_populates_fields => ("package main; import \"fmt\"; import \"encoding/json\"; type Person struct { Name string; Age int }; func main() { var p Person; json.Unmarshal([]byte(\"{\\\"Name\\\":\\\"Bob\\\",\\\"Age\\\":30}\"), &p); fmt.Println(p.Name); fmt.Println(p.Age) }", vec!["Bob", "30"]),
    unmarshal_struct_honors_json_tag => ("package main; import \"fmt\"; import \"encoding/json\"; type Item struct { ID int `json:\"id\"` }; func main() { var it Item; json.Unmarshal([]byte(\"{\\\"id\\\":99}\"), &it); fmt.Println(it.ID) }", vec!["99"]),
    unmarshal_slice_of_ints => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []int; json.Unmarshal([]byte(\"[10,20,30]\"), &s); fmt.Println(len(s)); fmt.Println(s[1]) }", vec!["3", "20"]),
    unmarshal_map_lookup => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { var m map[string]int; json.Unmarshal([]byte(\"{\\\"key\\\":7}\"), &m); fmt.Println(m[\"key\"]) }", vec!["7"]),
    unmarshal_nested_struct => ("package main; import \"fmt\"; import \"encoding/json\"; type Child struct { N int }; type Parent struct { Child Child }; func main() { var p Parent; json.Unmarshal([]byte(\"{\\\"Child\\\":{\\\"N\\\":5}}\"), &p); fmt.Println(p.Child.N) }", vec!["5"]),
    marshal_roundtrip_preserves_int => ("package main; import \"fmt\"; import \"encoding/json\"; func main() { orig := 123; b, _ := json.Marshal(orig); var back int; json.Unmarshal(b, &back); fmt.Println(back) }", vec!["123"]),
}

go_compile_cases! {
    json_marshal_indent_prefix => "package main; import \"encoding/json\"; func main() { _, _ = json.MarshalIndent(map[string]int{\"a\": 1}, \"\", \"  \") }",
    json_marshal_unexported_field_omitted => "package main; import \"encoding/json\"; type T struct { pub int; priv int }; func main() { _, _ = json.Marshal(T{pub: 1, priv: 2}) }",
    unmarshal_ignores_unknown_fields => "package main; import \"encoding/json\"; type S struct { X int }; func main() { var s S; _ = json.Unmarshal([]byte(\"{\\\"X\\\":1,\\\"extra\\\":\\\"y\\\"}\"), &s) }",
    unmarshal_to_interface_value => "package main; import \"encoding/json\"; func main() { var v interface{}; _ = json.Unmarshal([]byte(\"{\\\"n\\\":1}\"), &v) }",
    json_raw_message_holder => "package main; import \"encoding/json\"; func main() { var raw json.RawMessage; _ = json.Unmarshal([]byte(\"[1,2]\"), &raw) }",
    marshal_fixed_size_array => "package main; import \"encoding/json\"; func main() { _, _ = json.Marshal([2]int{4, 5}) }",
    unmarshal_empty_json_array => "package main; import \"encoding/json\"; func main() { var s []string; _ = json.Unmarshal([]byte(\"[]\"), &s) }",
    json_string_tag_int_field => "package main; import \"encoding/json\"; type N struct { Val int `json:\",string\"` }; func main() { _, _ = json.Marshal(N{Val: 7}) }",
    marshal_map_string_keys_only => "package main; import \"encoding/json\"; func main() { _, _ = json.Marshal(map[string]string{\"k\": \"v\"}) }",
    unmarshal_pointer_field_from_number => "package main; import \"encoding/json\"; type Box struct { N *int }; func main() { var b Box; _ = json.Unmarshal([]byte(\"{\\\"N\\\":8}\"), &b) }",
}
