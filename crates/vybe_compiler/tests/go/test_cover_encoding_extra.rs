//! encoding/xml, csv, gob, pem, ascii85, base32 - extra APIs not in
//! test_stdlib_encoding_misc.rs; one distinct API per compile smoke.


go_compile_cases! {
    // encoding/xml - beyond Marshal/Unmarshal
    xml_marshal_indent => "package main; import \"encoding/xml\"; type T struct { X int `xml:\"x\"` }; func main() { _, _ = xml.MarshalIndent(T{X: 1}, \"\", \"  \") }",
    xml_new_encoder => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { _ = xml.NewEncoder(bytes.NewBuffer(nil)) }",
    xml_new_decoder => "package main; import \"encoding/xml\"; import \"strings\"; func main() { _ = xml.NewDecoder(strings.NewReader(\"<T/>\")) }",
    xml_encoder_encode => "package main; import \"encoding/xml\"; import \"bytes\"; type T struct { X int `xml:\"x\"` }; func main() { e := xml.NewEncoder(bytes.NewBuffer(nil)); _ = e.Encode(T{X: 2}) }",
    xml_decoder_decode => "package main; import \"encoding/xml\"; import \"strings\"; type T struct { X int `xml:\"x\"` }; func main() { d := xml.NewDecoder(strings.NewReader(\"<T x=\\\"3\\\"></T>\")); var t T; _ = d.Decode(&t) }",
    xml_encoder_indent => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { e := xml.NewEncoder(bytes.NewBuffer(nil)); e.Indent(\"\", \"  \") }",
    xml_encoder_flush => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { e := xml.NewEncoder(bytes.NewBuffer(nil)); _ = e.Flush() }",
    xml_encoder_close => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { e := xml.NewEncoder(bytes.NewBuffer(nil)); _ = e.Close() }",
    xml_decoder_token => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _, _ = d.Token() }",
    xml_decoder_skip => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a><b/></a>\")); _ = d.Skip() }",
    xml_copy => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { dst := bytes.NewBuffer(nil); _ = xml.Copy(dst, bytes.NewBufferString(\"<a/>\")) }",
    xml_escape => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { _ = xml.Escape(bytes.NewBuffer(nil), []byte(\"<tag>\")) }",
    xml_escape_text => "package main; import \"encoding/xml\"; import \"bytes\"; func main() { _ = xml.EscapeText(bytes.NewBuffer(nil), []byte(\"text\")) }",
    xml_unescape => "package main; import \"encoding/xml\"; func main() { b, _ := xml.Unescape([]byte(\"&lt;a&gt;\")); _ = b }",
    xml_char_data => "package main; import \"encoding/xml\"; func main() { _ = xml.CharData([]byte(\"text\")) }",
    xml_header => "package main; import \"encoding/xml\"; func main() { _ = xml.Header }",

    // encoding/csv - beyond NewReader/NewWriter
    csv_reader_read => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\"a,b\")); _, _ = r.Read() }",
    csv_reader_read_all => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\"a,b\\nc,d\")); _, _ = r.ReadAll() }",
    csv_reader_fields_per_record => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\"a,b\")); r.FieldsPerRecord = 2; _, _ = r.Read() }",
    csv_reader_reuse_record => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\"a,b\")); r.ReuseRecord = true; _, _ = r.Read() }",
    csv_reader_lazy_quotes => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(`\"a\",b`)); r.LazyQuotes = true; _, _ = r.Read() }",
    csv_reader_trim_leading_space => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\" a , b \")); r.TrimLeadingSpace = true; _, _ = r.Read() }",
    csv_writer_write => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); _ = w.Write([]string{\"a\", \"b\"}) }",
    csv_writer_write_all => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); _ = w.WriteAll([][]string{{\"a\", \"b\"}, {\"c\", \"d\"}}) }",
    csv_writer_flush => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); w.Flush() }",
    csv_writer_error => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); _ = w.Error() }",
    csv_writer_comma => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); w.Comma = ';'; _ = w.Write([]string{\"a\", \"b\"}) }",

    // encoding/gob - beyond NewEncoder/NewDecoder
    gob_register => "package main; import \"encoding/gob\"; func main() { gob.Register(int(0)) }",
    gob_register_name => "package main; import \"encoding/gob\"; func main() { gob.RegisterName(\"Int\", int(0)) }",
    gob_encoder_encode => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { e := gob.NewEncoder(bytes.NewBuffer(nil)); _ = e.Encode(42) }",
    gob_decoder_decode => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { var v int; d := gob.NewDecoder(bytes.NewBuffer(nil)); _ = d.Decode(&v) }",

    // encoding/pem - beyond EncodeToMemory/Decode
    xml_decoder_entity => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); d.Entity = map[string]string{\"amp\": \"&\"}; _ = d.Entity[\"amp\"] }",
    pem_get_line => "package main; import \"encoding/pem\"; func main() { line, rest := pem.GetLine([]byte(\"TYPE rest\")); _, _ = line, rest }",
    pem_line_break => "package main; import \"encoding/pem\"; func main() { _ = pem.LineBreak }",

    // encoding/ascii85 - beyond NewEncoder
    ascii85_new_decoder => "package main; import \"encoding/ascii85\"; import \"bytes\"; func main() { _ = ascii85.NewDecoder(bytes.NewBufferString(\"<~00~>\")) }",
    ascii85_encode => "package main; import \"encoding/ascii85\"; func main() { dst := make([]byte, ascii85.MaxEncodedLen(4)); _ = ascii85.Encode(dst, []byte(\"go\")) }",
    ascii85_decode => "package main; import \"encoding/ascii85\"; func main() { dst := make([]byte, 4); _, _, _ = ascii85.Decode(dst, []byte(\"<~00~>\"), true) }",
    ascii85_max_encoded_len => "package main; import \"encoding/ascii85\"; func main() { _ = ascii85.MaxEncodedLen(8) }",
    ascii85_alphabet_len => "package main; import \"encoding/ascii85\"; func main() { _ = len(ascii85.Encode(make([]byte, 4), []byte(\"go\"))) }",

    // encoding/base32 - beyond StdEncoding.EncodeToString
    base32_decode_string => "package main; import \"encoding/base32\"; func main() { _, _ = base32.StdEncoding.DecodeString(\"MZXW6Y==\") }",
    base32_hex_encoding => "package main; import \"encoding/base32\"; func main() { _ = base32.HexEncoding.EncodeToString([]byte(\"go\")) }",
    base32_new_encoder => "package main; import \"encoding/base32\"; import \"bytes\"; func main() { _ = base32.NewEncoder(base32.StdEncoding, bytes.NewBuffer(nil)) }",
    base32_new_decoder => "package main; import \"encoding/base32\"; import \"strings\"; func main() { _ = base32.NewDecoder(base32.StdEncoding, strings.NewReader(\"MZXW6Y==\")) }",
    base32_encoding_with_padding => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.WithPadding('=') }",
    base32_encoding_encoded_len => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.EncodedLen(5) }",
    base32_encoding_decoded_len => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.DecodedLen(8) }",
    base32_encoding_strict => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.Strict() }",
    base32_no_padding => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.WithPadding(base32.NoPadding) }",
    base32_hex_decode_string => "package main; import \"encoding/base32\"; func main() { _, _ = base32.HexEncoding.DecodeString(\"CPNMU==\") }",
    base32_std_encoding_decode => "package main; import \"encoding/base32\"; func main() { dst := make([]byte, 8); _, _ = base32.StdEncoding.Decode(dst, []byte(\"MZXW6Y==\")) }",

    // encoding/xml - decoder/encoder extras
    xml_decoder_input_offset => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _ = d.InputOffset() }",
    xml_decoder_input_pos => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _, _ = d.InputPos() }",
    xml_decoder_raw_token => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _, _ = d.RawToken() }",
    xml_comment_type => "package main; import \"encoding/xml\"; func main() { _ = xml.Comment(\"note\") }",
    xml_directive_type => "package main; import \"encoding/xml\"; func main() { _ = xml.Directive(\"go\") }",
    xml_proc_inst_type => "package main; import \"encoding/xml\"; func main() { p := xml.ProcInst{}; _ = p.Target }",
    xml_decoder_raw_path => "package main; import \"encoding/xml\"; import \"strings\"; func main() { d := xml.NewDecoder(strings.NewReader(\"<a/>\")); _ = d.RawPath() }",

    // encoding/csv - constants and writer options
    csv_err_field_count => "package main; import \"encoding/csv\"; func main() { _ = csv.ErrFieldCount }",
    csv_writer_use_crlf => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { w := csv.NewWriter(bytes.NewBuffer(nil)); w.UseCRLF = true; _ = w.Write([]string{\"a\"}) }",
    csv_reader_comment => "package main; import \"encoding/csv\"; import \"strings\"; func main() { r := csv.NewReader(strings.NewReader(\"#note\\na,b\")); r.Comment = '#'; _, _ = r.Read() }",

    // encoding/gob - decoder/encoder methods
    gob_encoder_encode_value => "package main; import \"encoding/gob\"; import \"bytes\"; import \"reflect\"; func main() { e := gob.NewEncoder(bytes.NewBuffer(nil)); _ = e.EncodeValue(reflect.ValueOf(1)) }",
    gob_decoder_decode_value => "package main; import \"encoding/gob\"; import \"bytes\"; import \"reflect\"; func main() { d := gob.NewDecoder(bytes.NewBuffer(nil)); _ = d.DecodeValue(reflect.ValueOf(new(int)).Elem()) }",

    // encoding/ascii85 - stream helpers
    ascii85_new_encoder => "package main; import \"encoding/ascii85\"; import \"bytes\"; func main() { _ = ascii85.NewEncoder(bytes.NewBuffer(nil)) }",
    ascii85_encoder_write => "package main; import \"encoding/ascii85\"; import \"bytes\"; func main() { e := ascii85.NewEncoder(bytes.NewBuffer(nil)); _, _ = e.Write([]byte(\"go\")) }",
    ascii85_decoder_read => "package main; import \"encoding/ascii85\"; import \"bytes\"; func main() { r := ascii85.NewDecoder(bytes.NewBufferString(\"<~00~>\")); buf := make([]byte, 4); _, _ = r.Read(buf) }",
}
