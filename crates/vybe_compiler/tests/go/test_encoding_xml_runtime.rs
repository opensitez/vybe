//! encoding/xml runtime: Marshal/Unmarshal, attributes, nesting, xml.Name, CharData,
//! Encoder indent, decoder token loop, omitempty — distinct from compile-only smokes
//! in `test_cover_encoding_extra.rs` and `test_stdlib_encoding_misc.rs`.


go_run_cases! {
    xml_marshal_int_field_element => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n\"` }; func main() { b, _ := xml.Marshal(T{N: 7}); fmt.Println(string(b)) }",
        vec!["<T><n>7</n></T>"]
    ),
    xml_marshal_string_field_element => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"s\"` }; func main() { b, _ := xml.Marshal(T{S: \"go\"}); fmt.Println(string(b)) }",
        vec!["<T><s>go</s></T>"]
    ),
    xml_marshal_bool_true_element => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Ok bool `xml:\"ok\"` }; func main() { b, _ := xml.Marshal(T{Ok: true}); fmt.Println(string(b)) }",
        vec!["<T><ok>true</ok></T>"]
    ),
    xml_marshal_bool_false_element => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Ok bool `xml:\"ok\"` }; func main() { b, _ := xml.Marshal(T{Ok: false}); fmt.Println(string(b)) }",
        vec!["<T><ok>false</ok></T>"]
    ),
    xml_marshal_int_attribute => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,attr\"` }; func main() { b, _ := xml.Marshal(T{N: 42}); fmt.Println(string(b)) }",
        vec!["<T n=\"42\"></T>"]
    ),
    xml_marshal_string_attribute => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"label,attr\"` }; func main() { b, _ := xml.Marshal(T{S: \"vybe\"}); fmt.Println(string(b)) }",
        vec!["<T label=\"vybe\"></T>"]
    ),
    xml_unmarshal_int_attribute => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,attr\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T n=\"9\"/>`), &t); fmt.Println(t.N) }",
        vec!["9"]
    ),
    xml_unmarshal_string_attribute => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"name,attr\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T name=\"alice\"/>`), &t); fmt.Println(t.S) }",
        vec!["alice"]
    ),
    xml_unmarshal_element_text_to_int => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"count\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><count>15</count></T>`), &t); fmt.Println(t.N) }",
        vec!["15"]
    ),
    xml_unmarshal_element_text_to_string => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"msg\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><msg>hello</msg></T>`), &t); fmt.Println(t.S) }",
        vec!["hello"]
    ),
    xml_marshal_nested_child_struct => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type Inner struct { V int `xml:\"v\"` }; type Outer struct { Inner Inner `xml:\"inner\"` }; func main() { b, _ := xml.Marshal(Outer{Inner: Inner{V: 3}}); fmt.Println(string(b)) }",
        vec!["<Outer><inner><v>3</v></inner></Outer>"]
    ),
    xml_unmarshal_nested_child_struct => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type Inner struct { V int `xml:\"v\"` }; type Outer struct { Inner Inner `xml:\"inner\"` }; func main() { var o Outer; xml.Unmarshal([]byte(`<Outer><inner><v>8</inner></Outer>`), &o); fmt.Println(o.Inner.V) }",
        vec!["8"]
    ),
    xml_marshal_omitempty_skips_zero_int => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,omitempty\"` }; func main() { b, _ := xml.Marshal(T{}); fmt.Println(string(b)) }",
        vec!["<T></T>"]
    ),
    xml_marshal_omitempty_includes_nonzero_int => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,omitempty\"` }; func main() { b, _ := xml.Marshal(T{N: 1}); fmt.Println(string(b)) }",
        vec!["<T><n>1</n></T>"]
    ),
    xml_marshal_omitempty_skips_empty_string => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"s,omitempty\"` }; func main() { b, _ := xml.Marshal(T{}); fmt.Println(string(b)) }",
        vec!["<T></T>"]
    ),
    xml_marshal_omitempty_includes_nonempty_string => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { S string `xml:\"s,omitempty\"` }; func main() { b, _ := xml.Marshal(T{S: \"x\"}); fmt.Println(string(b)) }",
        vec!["<T><s>x</s></T>"]
    ),
    xml_name_local_field => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N xml.Name `xml:\"item\"` }; func main() { t := T{N: xml.Name{Local: \"widget\"}}; b, _ := xml.Marshal(t); fmt.Println(string(b)) }",
        vec!["<T><item>widget</item></T>"]
    ),
    xml_name_space_and_local_unmarshal => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N xml.Name `xml:\"tag\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><tag xmlns=\"urn:ex\">leaf</tag></T>`), &t); fmt.Println(t.N.Local); fmt.Println(t.N.Space) }",
        vec!["leaf", "urn:ex"]
    ),
    xml_char_data_string_conversion => (
        "package main; import \"fmt\"; import \"encoding/xml\"; func main() { cd := xml.CharData([]byte(\"payload\")); fmt.Println(string(cd)) }",
        vec!["payload"]
    ),
    xml_char_data_empty_slice => (
        "package main; import \"fmt\"; import \"encoding/xml\"; func main() { cd := xml.CharData([]byte{}); fmt.Println(len(cd)) }",
        vec!["0"]
    ),
    xml_marshal_chardata_innerxml => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Body string `xml:\",chardata\"` }; func main() { b, _ := xml.Marshal(T{Body: \"text\"}); fmt.Println(string(b)) }",
        vec!["<T>text</T>"]
    ),
    xml_unmarshal_chardata_innerxml => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Body string `xml:\",chardata\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T>inner</T>`), &t); fmt.Println(t.Body) }",
        vec!["inner"]
    ),
    xml_marshal_slice_repeated_elements => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Items []int `xml:\"item\"` }; func main() { b, _ := xml.Marshal(T{Items: []int{1, 2}}); fmt.Println(string(b)) }",
        vec!["<T><item>1</item><item>2</item></T>"]
    ),
    xml_unmarshal_slice_repeated_elements => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Items []int `xml:\"item\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><item>4</item><item>5</item></T>`), &t); fmt.Println(len(t.Items)); fmt.Println(t.Items[1]) }",
        vec!["2", "5"]
    ),
    xml_marshal_pointer_nil_omits => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { P *int `xml:\"p,omitempty\"` }; func main() { b, _ := xml.Marshal(T{}); fmt.Println(string(b)) }",
        vec!["<T></T>"]
    ),
    xml_marshal_pointer_nonnil_value => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { P *int `xml:\"p\"` }; func main() { n := 6; b, _ := xml.Marshal(T{P: &n}); fmt.Println(string(b)) }",
        vec!["<T><p>6</p></T>"]
    ),
    xml_unmarshal_pointer_allocates => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { P *int `xml:\"p\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><p>11</p></T>`), &t); fmt.Println(*t.P) }",
        vec!["11"]
    ),
    xml_marshal_float_field => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { F float64 `xml:\"f\"` }; func main() { b, _ := xml.Marshal(T{F: 2.5}); fmt.Println(string(b)) }",
        vec!["<T><f>2.5</f></T>"]
    ),
    xml_unmarshal_float_field => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { F float64 `xml:\"f\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><f>3.14</f></T>`), &t); fmt.Println(t.F) }",
        vec!["3.14"]
    ),
    xml_marshal_rename_with_tag => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Val int `xml:\"value\"` }; func main() { b, _ := xml.Marshal(T{Val: 99}); fmt.Println(string(b)) }",
        vec!["<T><value>99</value></T>"]
    ),
    xml_unmarshal_rename_with_tag => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Val int `xml:\"value\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><value>99</value></T>`), &t); fmt.Println(t.Val) }",
        vec!["99"]
    ),
    xml_marshal_dash_tag_omits_field => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { Hidden string `xml:\"-\"`; Pub int `xml:\"pub\"` }; func main() { b, _ := xml.Marshal(T{Hidden: \"secret\", Pub: 2}); fmt.Println(string(b)) }",
        vec!["<T><pub>2</pub></T>"]
    ),
    xml_marshal_indent_adds_newlines => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n\"` }; func main() { b, _ := xml.MarshalIndent(T{N: 1}, \"\", \"  \"); fmt.Println(len(b) > len([]byte(\"<T><n>1</n></T>\"))) }",
        vec!["true"]
    ),
    xml_marshal_indent_prefix => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n\"` }; func main() { b, _ := xml.MarshalIndent(T{N: 1}, \"--\", \"  \"); s := string(b); fmt.Println(s[0:2]) }",
        vec!["--"]
    ),
    xml_marshal_roundtrip_int => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,attr\"` }; func main() { orig := T{N: 55}; b, _ := xml.Marshal(orig); var back T; xml.Unmarshal(b, &back); fmt.Println(back.N) }",
        vec!["55"]
    ),
    xml_unmarshal_self_closing_tag => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,attr\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T n=\"3\"/>`), &t); fmt.Println(t.N) }",
        vec!["3"]
    ),
    xml_unmarshal_explicit_close_tag => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n,attr\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T n=\"4\"></T>`), &t); fmt.Println(t.N) }",
        vec!["4"]
    ),
    xml_escape_ampersand_in_text => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"bytes\"; func main() { var buf bytes.Buffer; xml.EscapeText(&buf, []byte(\"a&b\")); fmt.Println(buf.String()) }",
        vec!["a&amp;b"]
    ),
    xml_escape_less_than_in_text => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"bytes\"; func main() { var buf bytes.Buffer; xml.EscapeText(&buf, []byte(\"a<b\")); fmt.Println(buf.String()) }",
        vec!["a&lt;b"]
    ),
    xml_unescape_entity => (
        "package main; import \"fmt\"; import \"encoding/xml\"; func main() { b, _ := xml.Unescape([]byte(\"&lt;tag&gt;\")); fmt.Println(string(b)) }",
        vec!["<tag>"]
    ),
    xml_header_constant_non_empty => (
        "package main; import \"fmt\"; import \"encoding/xml\"; func main() { fmt.Println(len(xml.Header) > 0) }",
        vec!["true"]
    ),
    xml_decoder_token_start_element_local => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(`<root><child/></root>`)); tok, _ := d.Token(); start := tok.(xml.StartElement); fmt.Println(start.Name.Local) }",
        vec!["root"]
    ),
    xml_decoder_token_char_data => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(`<a>xy</a>`)); d.Token(); tok, _ := d.Token(); cd := tok.(xml.CharData); fmt.Println(string(cd)) }",
        vec!["xy"]
    ),
    xml_decoder_token_end_element => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(`<a/>`)); d.Token(); tok, _ := d.Token(); end := tok.(xml.EndElement); fmt.Println(end.Name.Local) }",
        vec!["a"]
    ),
    xml_decoder_token_loop_count => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(`<a><b/><c/></a>`)); n := 0; for { tok, err := d.Token(); if err != nil { break }; if _, ok := tok.(xml.StartElement); ok { n++ } }; fmt.Println(n) }",
        vec!["3"]
    ),
    xml_encoder_indent_then_encode => (
        "package main; import \"fmt\"; import \"encoding/xml\"; import \"bytes\"; type T struct { N int `xml:\"n\"` }; func main() { buf := bytes.NewBuffer(nil); e := xml.NewEncoder(buf); e.Indent(\"\", \"  \"); e.Encode(T{N: 2}); fmt.Println(buf.Len() > 0) }",
        vec!["true"]
    ),
    xml_marshal_two_attributes => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { A int `xml:\"a,attr\"`; B string `xml:\"b,attr\"` }; func main() { b, _ := xml.Marshal(T{A: 1, B: \"z\"}); s := string(b); fmt.Println(len(s) > 10) }",
        vec!["true"]
    ),
    xml_unmarshal_two_attributes => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { A int `xml:\"a,attr\"`; B string `xml:\"b,attr\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T a=\"2\" b=\"y\"/>`), &t); fmt.Println(t.A); fmt.Println(t.B) }",
        vec!["2", "y"]
    ),
    xml_marshal_empty_struct => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct {}; func main() { b, _ := xml.Marshal(T{}); fmt.Println(string(b)) }",
        vec!["<T></T>"]
    ),
    xml_unmarshal_missing_field_stays_zero => (
        "package main; import \"fmt\"; import \"encoding/xml\"; type T struct { N int `xml:\"n\"`; S string `xml:\"s\"` }; func main() { var t T; xml.Unmarshal([]byte(`<T><n>1</n></T>`), &t); fmt.Println(t.N); fmt.Println(t.S) }",
        vec!["1", ""]
    ),
}

go_compile_cases! {
    xml_encoder_decode_roundtrip_compile => "package main; import \"encoding/xml\"; import \"bytes\"; type T struct { X int `xml:\"x\"` }; func main() { buf := bytes.NewBuffer(nil); e := xml.NewEncoder(buf); e.Encode(T{X: 1}); d := xml.NewDecoder(buf); var t T; _ = d.Decode(&t) }",
    xml_decoder_raw_token_loop => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a><b/></a>\")); for { _, err := d.RawToken(); if err != nil { break } } }",
    xml_decoder_skip_nested => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a><b><c/></b></a>\")); d.Token(); _ = d.Skip() }",
    xml_copy_stream => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { dst := bytes.NewBuffer(nil); _ = xml.Copy(dst, bytes.NewBufferString(\"<root/>\")) }",
    xml_comment_type_compile => "package main; import \"encoding/xml\"; func main() { _ = xml.Comment(\"note\") }",
    xml_directive_type_compile => "package main; import \"encoding/xml\"; func main() { _ = xml.Directive(\"go\") }",
    xml_proc_inst_target_field => "package main; import \"encoding/xml\"; func main() { p := xml.ProcInst{Target: \"xml\"}; _ = p.Target }",
    xml_decoder_entity_map => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); d.Entity = map[string]string{\"copy\": \"©\"}; _ = d.Entity }",
    xml_decoder_input_offset_after_read => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _, _ = d.Token(); _ = d.InputOffset() }",
    xml_decoder_input_pos => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _, _ = d.InputPos() }",
}
