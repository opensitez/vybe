crate::js_cases! {
    textencoder_encode_ascii_length => {
        r#"
const bytes = new TextEncoder().encode("hello");
console.log(bytes.length);
"#,
        ["5"]
    };

    textencoder_encode_ascii_bytes => {
        r#"
const bytes = new TextEncoder().encode("ABC");
console.log(Array.from(bytes).join(","));
"#,
        ["65,66,67"]
    };

    textencoder_encode_empty_string => {
        r#"
const bytes = new TextEncoder().encode("");
console.log(bytes.length);
"#,
        ["0"]
    };

    textencoder_encode_latin1_uses_utf8_bytes => {
        r#"
const bytes = new TextEncoder().encode("é");
console.log(Array.from(bytes).join(","));
"#,
        ["195,169"]
    };

    textencoder_encode_astral_symbol_uses_four_bytes => {
        r#"
const bytes = new TextEncoder().encode("😀");
console.log(bytes.length);
"#,
        ["4"]
    };

    textencoder_returns_uint8array => {
        r#"
const bytes = new TextEncoder().encode("ok");
console.log(bytes instanceof Uint8Array);
"#,
        ["true"]
    };

    textdecoder_default_encoding_is_utf8 => {
        r#"
const decoder = new TextDecoder();
console.log(decoder.encoding);
"#,
        ["utf-8"]
    };

    textdecoder_decode_ascii_bytes => {
        r#"
const bytes = new Uint8Array([72, 101, 108, 108, 111]);
console.log(new TextDecoder().decode(bytes));
"#,
        ["Hello"]
    };

    textdecoder_decode_utf8_latin1_bytes => {
        r#"
const bytes = new Uint8Array([195, 169]);
console.log(new TextDecoder().decode(bytes));
"#,
        ["é"]
    };

    textdecoder_decode_utf8_astral_bytes => {
        r#"
const bytes = new Uint8Array([240, 159, 152, 128]);
const value = new TextDecoder().decode(bytes);
console.log(value);
console.log(value.length);
"#,
        ["😀", "2"]
    };

    textdecoder_decode_empty_array => {
        r#"
console.log(new TextDecoder().decode(new Uint8Array()) === "");
"#,
        ["true"]
    };

    textdecoder_decode_subarray_view => {
        r#"
const bytes = new Uint8Array([88, 72, 105, 89]);
console.log(new TextDecoder().decode(bytes.subarray(1, 3)));
"#,
        ["Hi"]
    };

    textdecoder_decode_from_arraybuffer => {
        r#"
const bytes = new Uint8Array([79, 75]);
console.log(new TextDecoder().decode(bytes.buffer));
"#,
        ["OK"]
    };

    textdecoder_decode_partial_view_only_reads_view => {
        r#"
const bytes = new Uint8Array([65, 66, 67, 68]);
const view = new Uint8Array(bytes.buffer, 1, 2);
console.log(new TextDecoder().decode(view));
"#,
        ["BC"]
    };

    textencoder_encodeinto_reports_read_and_written => {
        r#"
const encoder = new TextEncoder();
const dest = new Uint8Array(10);
const result = encoder.encodeInto("Hi", dest);
console.log(result.read);
console.log(result.written);
console.log(Array.from(dest.slice(0, 2)).join(","));
"#,
        ["2", "2", "72,105"]
    };

    textdecoder_fatal_option_property => {
        r#"
const decoder = new TextDecoder("utf-8", { fatal: true });
console.log(decoder.fatal);
"#,
        ["true"]
    };

    textdecoder_ignorebom_option_property => {
        r#"
const decoder = new TextDecoder("utf-8", { ignoreBOM: true });
console.log(decoder.ignoreBOM);
"#,
        ["true"]
    };

    btoa_encodes_ascii_text => {
        r#"
console.log(btoa("Hello"));
"#,
        ["SGVsbG8="]
    };

    atob_decodes_ascii_text => {
        r#"
console.log(atob("SGVsbG8="));
"#,
        ["Hello"]
    };

    base64_roundtrip_with_spaces_and_punctuation => {
        r#"
const value = "Hello, world!";
console.log(atob(btoa(value)));
"#,
        ["Hello, world!"]
    };

    btoa_empty_string_is_empty => {
        r#"
console.log(btoa(""));
"#,
        [""]
    };

    atob_empty_string_is_empty => {
        r#"
console.log(atob(""));
"#,
        [""]
    };

    atob_decodes_plus_and_slash_alphabet => {
        r#"
console.log(atob("+/8=").length);
"#,
        ["2"]
    };

    atob_invalid_input_throws => {
        r#"
try {
  atob("***");
  console.log("no error");
} catch (error) {
  console.log(error instanceof Error);
}
"#,
        ["true"]
    };

    btoa_non_latin1_input_throws => {
        r#"
try {
  btoa("😀");
  console.log("no error");
} catch (error) {
  console.log(error instanceof Error);
}
"#,
        ["true"]
    };

    textdecoder_decode_bom_prefixed_utf8 => {
        r#"
const bytes = new Uint8Array([239, 187, 191, 65]);
console.log(new TextDecoder().decode(bytes));
"#,
        ["A"]
    };
}
