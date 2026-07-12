crate::js_cases! {
    textencoder_encodeinto_partial_buffer_reports_partial_write => {
        r#"
const encoder = new TextEncoder();
const dest = new Uint8Array(2);
const result = encoder.encodeInto("hello", dest);
console.log(result.read);
console.log(result.written);
"#,
        ["2", "2"]
    };
    textdecoder_invalid_utf8_uses_replacement_character => {
        r#"
const bytes = new Uint8Array([0xE2, 0x28, 0xA1]);
console.log(new TextDecoder().decode(bytes).includes("\uFFFD"));
"#,
        ["true"]
    };
    textdecoder_fatal_invalid_utf8_throws => {
        r#"
try {
  new TextDecoder("utf-8", { fatal: true }).decode(new Uint8Array([0xE2, 0x28, 0xA1]));
  console.log("no error");
} catch (error) {
  console.log(error instanceof Error);
}
"#,
        ["true"]
    };
    textdecoder_ignorebom_false_strips_utf8_bom => {
        r#"
const bytes = new Uint8Array([239, 187, 191, 65]);
console.log(new TextDecoder("utf-8", { ignoreBOM: false }).decode(bytes));
"#,
        ["A"]
    };
    textdecoder_ignorebom_true_preserves_utf8_bom_codepoint => {
        r#"
const bytes = new Uint8Array([239, 187, 191, 65]);
const value = new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
console.log(value.charCodeAt(0));
console.log(value.slice(1));
"#,
        ["65279", "A"]
    };
    atob_accepts_padded_base64 => { r#"console.log(atob("TQ=="));"#, ["M"] };
    atob_accepts_unpadded_single_byte_length_when_valid => { r#"console.log(atob("TWE=").length);"#, ["2"] };
    btoa_ascii_nul_byte_roundtrips => {
        r#"
const value = "A\u0000B";
console.log(atob(btoa(value)).length);
"#,
        ["3"]
    };
    textencoder_reencodes_decoded_utf8_roundtrip => {
        r#"
const s = new TextDecoder().decode(new Uint8Array([195, 169]));
console.log(Array.from(new TextEncoder().encode(s)).join(","));
"#,
        ["195,169"]
    };
    textdecoder_decode_uint8array_slice_not_full_buffer => {
        r#"
const arr = new Uint8Array([88, 89, 90]);
console.log(new TextDecoder().decode(arr.slice(1)));
"#,
        ["YZ"]
    };
    btoa_space_character_roundtrips => {
        r#"
console.log(atob(btoa(" ")) === " ");
"#,
        ["true"]
    };
    atob_invalid_padding_throws => {
        r#"
try {
  atob("=");
  console.log("no error");
} catch (error) {
  console.log(error instanceof Error);
}
"#,
        ["true"]
    };
    textdecoder_encoding_property_remains_utf8_when_specified => {
        r#"
console.log(new TextDecoder("utf-8").encoding);
"#,
        ["utf-8"]
    };
    textencoder_multiple_calls_are_deterministic => {
        r#"
const enc = new TextEncoder();
console.log(Array.from(enc.encode("ok")).join(",") === Array.from(enc.encode("ok")).join(","));
"#,
        ["true"]
    };
    textdecoder_multiple_calls_are_deterministic => {
        r#"
const dec = new TextDecoder();
console.log(dec.decode(new Uint8Array([79, 75])) === dec.decode(new Uint8Array([79, 75])));
"#,
        ["true"]
    };
    textencoder_utf8_length_for_two_accented_chars_is_four => {
        r#"
console.log(new TextEncoder().encode("éé").length);
"#,
        ["4"]
    };
    atob_roundtrip_with_slash_character => {
        r#"
const value = "a/b";
console.log(atob(btoa(value)));
"#,
        ["a/b"]
    };
    btoa_roundtrip_with_plus_character => {
        r#"
const value = "a+b";
console.log(atob(btoa(value)));
"#,
        ["a+b"]
    };
    textdecoder_empty_arraybuffer_decodes_empty_string => {
        r#"
console.log(new TextDecoder().decode(new ArrayBuffer(0)) === "");
"#,
        ["true"]
    };
    textencoder_encodeinto_zero_length_buffer_writes_nothing => {
        r#"
const result = new TextEncoder().encodeInto("hello", new Uint8Array(0));
console.log(result.written);
"#,
        ["0"]
    };
}
