//! encoding/hex and encoding/base64: Encode, Decode, Dump, and StdEncoding round-trips.

go_run_cases! {
    hex_encode_empty_slice => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { src := []byte{}; dst := make([]byte, hex.EncodedLen(len(src))); n := hex.Encode(dst, src); fmt.Println(n) }",
        vec!["0"]
    ),
    hex_encode_single_zero_byte => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { fmt.Println(hex.EncodeToString([]byte{0})) }",
        vec!["00"]
    ),
    hex_encode_ascii_word => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { src := []byte(\"go\"); dst := make([]byte, hex.EncodedLen(len(src))); hex.Encode(dst, src); fmt.Println(string(dst)) }",
        vec!["676f"]
    ),
    hex_encode_max_byte_ff => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { fmt.Println(hex.EncodeToString([]byte{0xff})) }",
        vec!["ff"]
    ),
    hex_encode_to_string_pair => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { fmt.Println(hex.EncodeToString([]byte{0x0a, 0x0b})) }",
        vec!["0a0b"]
    ),
    hex_decode_lowercase_pair => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { b, err := hex.DecodeString(\"6162\"); fmt.Println(string(b)); fmt.Println(err == nil) }",
        vec!["ab", "true"]
    ),
    hex_decode_uppercase_accepted => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { b, err := hex.DecodeString(\"4142\"); fmt.Println(string(b)); fmt.Println(err == nil) }",
        vec!["AB", "true"]
    ),
    hex_decode_empty_string => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { b, err := hex.DecodeString(\"\"); fmt.Println(len(b)); fmt.Println(err == nil) }",
        vec!["0", "true"]
    ),
    hex_encode_decode_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { orig := []byte{1, 2, 3, 250}; enc := hex.EncodeToString(orig); back, _ := hex.DecodeString(enc); fmt.Println(len(back)); fmt.Println(int(back[3])) }",
        vec!["4", "250"]
    ),
    hex_decode_odd_length_errors => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { _, err := hex.DecodeString(\"414\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    hex_decode_invalid_char_errors => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { _, err := hex.DecodeString(\"gh\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    hex_dump_includes_offset_and_ascii => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { dump := string(hex.Dump([]byte(\"ab\"))); fmt.Println(len(dump) > 0); fmt.Println(dump[0:8]) }",
        vec!["true", "00000000"]
    ),
    hex_encoded_len_doubles_input => (
        "package main; import \"fmt\"; import \"encoding/hex\"; func main() { fmt.Println(hex.EncodedLen(5)) }",
        vec!["10"]
    ),
    base64_std_encode_empty => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { fmt.Println(base64.StdEncoding.EncodeToString([]byte{})) }",
        vec![""]
    ),
    base64_std_encode_single_byte_padding => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { fmt.Println(base64.StdEncoding.EncodeToString([]byte(\"f\"))) }",
        vec!["Zg=="]
    ),
    base64_std_encode_two_bytes_padding => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { fmt.Println(base64.StdEncoding.EncodeToString([]byte(\"fo\"))) }",
        vec!["Zm8="]
    ),
    base64_std_encode_three_bytes_no_padding => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { fmt.Println(base64.StdEncoding.EncodeToString([]byte(\"foo\"))) }",
        vec!["Zm9v"]
    ),
    base64_std_decode_without_padding => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { b, err := base64.StdEncoding.DecodeString(\"Zm9v\"); fmt.Println(string(b)); fmt.Println(err == nil) }",
        vec!["foo", "true"]
    ),
    base64_std_decode_with_padding => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { b, err := base64.StdEncoding.DecodeString(\"Zg==\"); fmt.Println(string(b)); fmt.Println(err == nil) }",
        vec!["f", "true"]
    ),
    base64_std_encode_decode_roundtrip => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { orig := []byte{0, 1, 255}; enc := base64.StdEncoding.EncodeToString(orig); back, _ := base64.StdEncoding.DecodeString(enc); fmt.Println(len(back)); fmt.Println(int(back[2])) }",
        vec!["3", "255"]
    ),
    base64_std_encoded_len_formula => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { fmt.Println(base64.StdEncoding.EncodedLen(1)); fmt.Println(base64.StdEncoding.EncodedLen(3)) }",
        vec!["4", "4"]
    ),
    base64_std_decode_empty_string => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { b, err := base64.StdEncoding.DecodeString(\"\"); fmt.Println(len(b)); fmt.Println(err == nil) }",
        vec!["0", "true"]
    ),
    base64_std_decode_invalid_char_errors => (
        "package main; import \"fmt\"; import \"encoding/base64\"; func main() { _, err := base64.StdEncoding.DecodeString(\"!!!\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    hex_decode_into_dst_buffer => "package main; import \"encoding/hex\"; func main() { dst := make([]byte, 2); _, _ = hex.Decode(dst, []byte(\"6162\")) }",
    hex_decoded_len_half_encoded => "package main; import \"encoding/hex\"; func main() { _ = hex.DecodedLen(4) }",
    hex_dumper_writer => "package main; import \"bytes\"; import \"encoding/hex\"; func main() { var buf bytes.Buffer; w := hex.Dumper(&buf); _, _ = w.Write([]byte(\"x\")); w.Close() }",
    hex_invalid_byte_error_value => "package main; import \"encoding/hex\"; func main() { _ = hex.InvalidByte }",
    hex_append_encode_slice => "package main; import \"encoding/hex\"; func main() { b := []byte{}; _ = hex.AppendEncode(b, []byte(\"a\")) }",
    base64_std_decode_into_dst_buffer => "package main; import \"encoding/base64\"; func main() { dst := make([]byte, 4); _, _ = base64.StdEncoding.Decode(dst, []byte(\"Zm9v\")) }",
    base64_std_decoded_len_estimate => "package main; import \"encoding/base64\"; func main() { _ = base64.StdEncoding.DecodedLen(4) }",
    base64_raw_std_encoding_no_padding => "package main; import \"encoding/base64\"; func main() { _ = base64.RawStdEncoding.EncodeToString([]byte(\"f\")) }",
    base64_url_encoding_alphabet => "package main; import \"encoding/base64\"; func main() { _ = base64.URLEncoding.EncodeToString([]byte(\"?\")) }",
    base64_with_padding_custom => "package main; import \"encoding/base64\"; func main() { enc := base64.StdEncoding.WithPadding('*'); _ = enc.EncodeToString([]byte(\"f\")) }",
}
