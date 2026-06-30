//! encoding/json advanced Unmarshal/Marshal: embedded fields, string tags, null pointer
//! slice elements, unicode escapes, MarshalIndent, RawMessage, DisallowUnknownFields, UseNumber
//! — distinct from `test_json_marshal.rs`.

use crate::helpers::*;

go_run_cases! {
    json_unmarshal_embedded_promoted_field => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Inner struct { N int }; type Outer struct { Inner }; func main() { var o Outer; json.Unmarshal([]byte(`{\"N\":5}`), &o); fmt.Println(o.N) }",
        vec!["5"]
    ),
    json_unmarshal_embedded_nested_struct => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Addr struct { City string }; type Person struct { Name string; Addr }; func main() { var p Person; json.Unmarshal([]byte(`{\"Name\":\"Ann\",\"City\":\"Paris\"}`), &p); fmt.Println(p.Name); fmt.Println(p.Addr.City) }",
        vec!["Ann", "Paris"]
    ),
    json_unmarshal_embedded_anonymous_tag => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Meta struct { Ver int `json:\"ver\"` }; type Doc struct { Meta; Title string }; func main() { var d Doc; json.Unmarshal([]byte(`{\"ver\":2,\"Title\":\"x\"}`), &d); fmt.Println(d.Ver); fmt.Println(d.Title) }",
        vec!["2", "x"]
    ),
    json_unmarshal_embedded_pointer_nil => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Inner struct { N int }; type Outer struct { *Inner }; func main() { var o Outer; json.Unmarshal([]byte(`{\"N\":3}`), &o); fmt.Println(o.Inner.N) }",
        vec!["3"]
    ),
    json_unmarshal_embedded_shadowed_field => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Base struct { ID int `json:\"id\"` }; type Ext struct { Base; ID int `json:\"ext_id\"` }; func main() { var e Ext; json.Unmarshal([]byte(`{\"id\":1,\"ext_id\":9}`), &e); fmt.Println(e.Base.ID); fmt.Println(e.ID) }",
        vec!["1", "9"]
    ),
    json_unmarshal_string_tag_int_value => (
        "package main; import \"fmt\"; import \"encoding/json\"; type N struct { Val int `json:\",string\"` }; func main() { var n N; json.Unmarshal([]byte(`{\"Val\":\"42\"}`), &n); fmt.Println(n.Val) }",
        vec!["42"]
    ),
    json_unmarshal_string_tag_negative => (
        "package main; import \"fmt\"; import \"encoding/json\"; type N struct { Val int `json:\",string\"` }; func main() { var n N; json.Unmarshal([]byte(`{\"Val\":\"-7\"}`), &n); fmt.Println(n.Val) }",
        vec!["-7"]
    ),
    json_marshal_string_tag_int => (
        "package main; import \"fmt\"; import \"encoding/json\"; type N struct { Val int `json:\",string\"` }; func main() { b, _ := json.Marshal(N{Val: 12}); fmt.Println(string(b)) }",
        vec!["{\"Val\":\"12\"}"]
    ),
    json_marshal_string_tag_zero => (
        "package main; import \"fmt\"; import \"encoding/json\"; type N struct { Val int `json:\",string\"` }; func main() { b, _ := json.Marshal(N{}); fmt.Println(string(b)) }",
        vec!["{\"Val\":\"0\"}"]
    ),
    json_unmarshal_string_tag_bool => (
        "package main; import \"fmt\"; import \"encoding/json\"; type B struct { Ok bool `json:\",string\"` }; func main() { var b B; json.Unmarshal([]byte(`{\"Ok\":\"true\"}`), &b); fmt.Println(b.Ok) }",
        vec!["true"]
    ),
    json_unmarshal_null_pointer_slice_element => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []*int; json.Unmarshal([]byte(`[1,null,3]`), &s); fmt.Println(s[0] != nil); fmt.Println(s[1] == nil); fmt.Println(*s[2]) }",
        vec!["true", "true", "3"]
    ),
    json_unmarshal_null_pointer_struct_field => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Box struct { N *int }; func main() { var b Box; json.Unmarshal([]byte(`{\"N\":null}`), &b); fmt.Println(b.N == nil) }",
        vec!["true"]
    ),
    json_unmarshal_null_pointer_then_value => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Box struct { N *int }; func main() { var b Box; json.Unmarshal([]byte(`{\"N\":8}`), &b); fmt.Println(*b.N) }",
        vec!["8"]
    ),
    json_unmarshal_null_in_pointer_array => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []*string; json.Unmarshal([]byte(`[\"a\",null,\"c\"]`), &s); fmt.Println(*s[0]); fmt.Println(s[1] == nil); fmt.Println(*s[2]) }",
        vec!["a", "true", "c"]
    ),
    json_unmarshal_unicode_escape_basic => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s string; json.Unmarshal([]byte(`\"\\u0047\\u006f\"`), &s); fmt.Println(s) }",
        vec!["Go"]
    ),
    json_unmarshal_unicode_escape_cyrillic => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s string; json.Unmarshal([]byte(`\"\\u043f\\u0440\\u0438\\u0432\\u0435\\u0442\"`), &s); fmt.Println(len(s)) }",
        vec!["6"]
    ),
    json_unmarshal_unicode_surrogate_pair => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s string; json.Unmarshal([]byte(`\"\\uD83D\\uDE00\"`), &s); fmt.Println(len([]rune(s))) }",
        vec!["1"]
    ),
    json_marshal_unicode_non_ascii => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.Marshal(\"café\"); fmt.Println(string(b)) }",
        vec!["\"café\""]
    ),
    json_unmarshal_unicode_in_object_key => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var m map[string]int; json.Unmarshal([]byte(`{\"\\u006b\":1}`), &m); fmt.Println(m[\"k\"]) }",
        vec!["1"]
    ),
    json_marshal_indent_two_space_prefix => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.MarshalIndent(map[string]int{\"a\": 1}, \"\", \"  \"); s := string(b); fmt.Println(s[0:1] == \"{\") }",
        vec!["true"]
    ),
    json_marshal_indent_custom_prefix => (
        "package main; import \"fmt\"; import \"encoding/json\"; type T struct { N int }; func main() { b, _ := json.MarshalIndent(T{N: 1}, \">\", \"  \"); s := string(b); fmt.Println(s[0:1]) }",
        vec![">"]
    ),
    json_marshal_indent_nested_struct => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Inner struct { V int }; type Outer struct { Inner Inner }; func main() { b, _ := json.MarshalIndent(Outer{Inner: Inner{V: 2}}, \"\", \"  \"); fmt.Println(len(b) > 10) }",
        vec!["true"]
    ),
    json_marshal_indent_empty_prefix => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.MarshalIndent([]int{1}, \"\", \"\"); fmt.Println(len(b) > 0) }",
        vec!["true"]
    ),
    json_raw_message_envelope_object => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Envelope struct { Payload json.RawMessage }; func main() { var e Envelope; json.Unmarshal([]byte(`{\"Payload\":{\"x\":1}}`), &e); var m map[string]int; json.Unmarshal(e.Payload, &m); fmt.Println(m[\"x\"]) }",
        vec!["1"]
    ),
    json_raw_message_envelope_array => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Envelope struct { Data json.RawMessage }; func main() { var e Envelope; json.Unmarshal([]byte(`{\"Data\":[1,2]}`), &e); var s []int; json.Unmarshal(e.Data, &s); fmt.Println(len(s)); fmt.Println(s[1]) }",
        vec!["2", "2"]
    ),
    json_raw_message_marshal_passthrough => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { raw := json.RawMessage(`{\"k\":7}`); b, _ := json.Marshal(raw); fmt.Println(string(b)) }",
        vec!["{\"k\":7}"]
    ),
    json_raw_message_string_literal => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { raw := json.RawMessage(`\"hello\"`); var s string; json.Unmarshal(raw, &s); fmt.Println(s) }",
        vec!["hello"]
    ),
    json_unmarshal_embedded_slice => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Tags []string; type Post struct { Tags; Title string }; func main() { var p Post; json.Unmarshal([]byte(`{\"Title\":\"t\",\"Tags\":[\"a\",\"b\"]}`), &p); fmt.Println(len(p.Tags)); fmt.Println(p.Tags[1]) }",
        vec!["2", "b"]
    ),
    json_unmarshal_string_tag_float_as_string => (
        "package main; import \"fmt\"; import \"encoding/json\"; type F struct { Val float64 `json:\",string\"` }; func main() { var f F; json.Unmarshal([]byte(`{\"Val\":\"3.14\"}`), &f); fmt.Println(int(f.Val*100)) }",
        vec!["314"]
    ),
    json_unmarshal_pointer_to_struct_in_slice => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Item struct { N int }; func main() { var s []*Item; json.Unmarshal([]byte(`[{\"N\":1},{\"N\":2}]`), &s); fmt.Println(s[1].N) }",
        vec!["2"]
    ),
    json_unmarshal_embedded_time_layout => (
        "package main; import \"fmt\"; import \"encoding/json\"; type When struct { Year int `json:\"year\"` }; type Event struct { When; Name string }; func main() { var e Event; json.Unmarshal([]byte(`{\"year\":2024,\"Name\":\"launch\"}`), &e); fmt.Println(e.Year); fmt.Println(e.Name) }",
        vec!["2024", "launch"]
    ),
    json_unmarshal_unicode_escape_in_array => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []string; json.Unmarshal([]byte(`[\"\\u0061\",\"\\u0062\"]`), &s); fmt.Println(s[0]); fmt.Println(s[1]) }",
        vec!["a", "b"]
    ),
    json_marshal_indent_map_sorted => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.MarshalIndent(map[string]int{\"b\": 2, \"a\": 1}, \"\", \"  \"); fmt.Println(len(b) > 5) }",
        vec!["true"]
    ),
    json_unmarshal_null_slice_becomes_nil => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s []int; json.Unmarshal([]byte(`null`), &s); fmt.Println(s == nil) }",
        vec!["true"]
    ),
    json_unmarshal_empty_object_to_struct => (
        "package main; import \"fmt\"; import \"encoding/json\"; type S struct { X int; Y string }; func main() { var s S; json.Unmarshal([]byte(`{}`), &s); fmt.Println(s.X); fmt.Println(s.Y == \"\") }",
        vec!["0", "true"]
    ),
    json_unmarshal_string_tag_uint => (
        "package main; import \"fmt\"; import \"encoding/json\"; type U struct { Val uint `json:\",string\"` }; func main() { var u U; json.Unmarshal([]byte(`{\"Val\":\"255\"}`), &u); fmt.Println(u.Val) }",
        vec!["255"]
    ),
    json_unmarshal_embedded_map_type => (
        "package main; import \"fmt\"; import \"encoding/json\"; type Meta map[string]string; type Doc struct { Meta; ID int }; func main() { var d Doc; json.Unmarshal([]byte(`{\"ID\":1,\"k\":\"v\"}`), &d); fmt.Println(d.ID); fmt.Println(d.Meta[\"k\"]) }",
        vec!["1", "v"]
    ),
    json_raw_message_null_value => (
        "package main; import \"fmt\"; import \"encoding/json\"; type W struct { Raw json.RawMessage }; func main() { var w W; json.Unmarshal([]byte(`{\"Raw\":null}`), &w); fmt.Println(w.Raw == nil) }",
        vec!["true"]
    ),
    json_marshal_indent_slice_elements => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { b, _ := json.MarshalIndent([]int{1, 2}, \"  \", \"  \"); fmt.Println(len(b) > 0) }",
        vec!["true"]
    ),
    json_unmarshal_unicode_bmp_char => (
        "package main; import \"fmt\"; import \"encoding/json\"; func main() { var s string; json.Unmarshal([]byte(`\"\\u00e9\"`), &s); fmt.Println(s) }",
        vec!["é"]
    ),
}

go_compile_cases! {
    json_decoder_disallow_unknown_fields => "package main; import \"encoding/json\"; import \"bytes\"; type S struct { X int }; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{\"X\":1,\"extra\":2}`))); dec.DisallowUnknownFields(); _ = dec.Decode(&S{}) }",
    json_decoder_disallow_unknown_nested => "package main; import \"encoding/json\"; import \"bytes\"; type Inner struct { N int }; type Outer struct { Inner Inner }; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{\"Inner\":{\"N\":1,\"bad\":2}}`))); dec.DisallowUnknownFields(); var o Outer; _ = dec.Decode(&o) }",
    json_decoder_use_number_int => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`12345`))); dec.UseNumber(); var v interface{}; _ = dec.Decode(&v) }",
    json_decoder_use_number_object => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{\"n\":99}`))); dec.UseNumber(); var v map[string]interface{}; _ = dec.Decode(&v) }",
    json_decoder_use_number_array => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`[1,2,3]`))); dec.UseNumber(); var v []interface{}; _ = dec.Decode(&v) }",
    json_decoder_use_number_large => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`9999999999999999999`))); dec.UseNumber(); var v interface{}; _ = dec.Decode(&v) }",
    json_marshal_indent_string_prefix => "package main; import \"encoding/json\"; func main() { _, _ = json.MarshalIndent(struct{ A int }{1}, \"tab\", \"  \") }",
    json_marshal_indent_pointer => "package main; import \"encoding/json\"; func main() { n := 5; _, _ = json.MarshalIndent(&n, \"\", \"  \") }",
    json_raw_message_unmarshal_nested => "package main; import \"encoding/json\"; func main() { var raw json.RawMessage; _ = json.Unmarshal([]byte(`{\"a\":{\"b\":1}}`), &raw) }",
    json_unmarshal_embedded_interface => "package main; import \"encoding/json\"; type Base struct { N int }; type Ext struct { Base; Extra string }; func main() { var e Ext; _ = json.Unmarshal([]byte(`{\"N\":1,\"Extra\":\"x\"}`), &e) }",
    json_unmarshal_string_tag_int64 => "package main; import \"encoding/json\"; type T struct { V int64 `json:\",string\"` }; func main() { var t T; _ = json.Unmarshal([]byte(`{\"V\":\"9223372036854775807\"}`), &t) }",
    json_unmarshal_null_pointer_in_struct_slice => "package main; import \"encoding/json\"; type Item struct { N *int }; func main() { var s []Item; _ = json.Unmarshal([]byte(`[{\"N\":null},{\"N\":1}]`), &s) }",
    json_unmarshal_unicode_escape_in_key => "package main; import \"encoding/json\"; func main() { var m map[string]int; _ = json.Unmarshal([]byte(`{\"\\u0061\":1}`), &m) }",
    json_decoder_token_use_number => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{\"x\":1.5}`))); dec.UseNumber(); _, _ = dec.Token() }",
    json_marshal_raw_message_field => "package main; import \"encoding/json\"; type W struct { Raw json.RawMessage `json:\"raw\"` }; func main() { _, _ = json.Marshal(W{Raw: json.RawMessage(`{\"k\":1}`)}) }",
    json_unmarshal_embedded_anonymous_pointer => "package main; import \"encoding/json\"; type Inner struct { V int }; type Outer struct { *Inner }; func main() { var o Outer; _ = json.Unmarshal([]byte(`{\"V\":9}`), &o) }",
    json_unmarshal_string_tag_omitempty => "package main; import \"encoding/json\"; type T struct { V int `json:\",string,omitempty\"` }; func main() { _, _ = json.Marshal(T{}) }",
    json_decoder_disallow_unknown_array_elem => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`[1,{\"x\":1,\"y\":2}]`))); dec.DisallowUnknownFields(); var v []interface{}; _ = dec.Decode(&v) }",
    json_marshal_indent_bool => "package main; import \"encoding/json\"; func main() { _, _ = json.MarshalIndent(true, \"\", \"  \") }",
    json_unmarshal_embedded_two_levels => "package main; import \"encoding/json\"; type A struct { N int }; type B struct { A }; type C struct { B }; func main() { var c C; _ = json.Unmarshal([]byte(`{\"N\":4}`), &c) }",
    json_decoder_use_number_float_string => "package main; import \"encoding/json\"; import \"bytes\"; func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`\"3.14\"`))); dec.UseNumber(); var v interface{}; _ = dec.Decode(&v) }",
    json_raw_message_marshal_nil => "package main; import \"encoding/json\"; func main() { var raw json.RawMessage; _, _ = json.Marshal(raw) }",
    json_unmarshal_pointer_slice_all_null => "package main; import \"encoding/json\"; func main() { var s []*int; _ = json.Unmarshal([]byte(`[null,null]`), &s) }",
    json_marshal_indent_empty_map => "package main; import \"encoding/json\"; func main() { _, _ = json.MarshalIndent(map[string]int{}, \"\", \"  \") }",
    json_unmarshal_embedded_with_json_tag => "package main; import \"encoding/json\"; type Meta struct { Tag string `json:\"tag\"` }; type Doc struct { Meta; Body string }; func main() { var d Doc; _ = json.Unmarshal([]byte(`{\"tag\":\"v1\",\"Body\":\"hi\"}`), &d) }",
}
