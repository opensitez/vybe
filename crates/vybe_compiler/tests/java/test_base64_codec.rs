use crate::helpers::run_main;

#[test]
fn base64_encode_simple_string() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("Hi".getBytes()));"#);
    assert_eq!(out, vec!["SGk="]);
}

#[test]
fn base64_encode_empty_bytes() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString(new byte[0]));"#);
    assert_eq!(out, vec![""]);
}

#[test]
fn base64_encode_hello_world() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("Hello".getBytes()));"#);
    assert_eq!(out, vec!["SGVsbG8="]);
}

#[test]
fn base64_decode_roundtrip() {
    let out = run_main(r#"byte[] enc = java.util.Base64.getEncoder().encode("Vybe".getBytes()); String dec = new String(java.util.Base64.getDecoder().decode(enc)); System.out.println(dec);"#);
    assert_eq!(out, vec!["Vybe"]);
}

#[test]
fn base64_decode_known_value() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getDecoder().decode("SGVsbG8=")));"#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn base64_encode_without_padding() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().withoutPadding().encodeToString("fo".getBytes()));"#);
    assert_eq!(out, vec!["Zm8"]);
}

#[test]
fn base64_url_encode_no_padding() {
    let out = run_main(r#"System.out.println(java.util.Base64.getUrlEncoder().withoutPadding().encodeToString("a?b".getBytes()));"#);
    assert_eq!(out, vec!["YT9i"]);
}

#[test]
fn base64_url_decode_roundtrip() {
    let out = run_main(r#"String enc = java.util.Base64.getUrlEncoder().encodeToString("test/data".getBytes()); System.out.println(new String(java.util.Base64.getUrlDecoder().decode(enc)));"#);
    assert_eq!(out, vec!["test/data"]);
}

#[test]
fn base64_mime_encode_wraps() {
    let out = run_main(r#"System.out.println(java.util.Base64.getMimeEncoder().encodeToString("abcd".getBytes()).length() > 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_mime_decode() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getMimeDecoder().decode("SGVsbG8=")));"#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn base64_encode_single_byte_a() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("A".getBytes()));"#);
    assert_eq!(out, vec!["QQ=="]);
}

#[test]
fn base64_encode_two_bytes_ab() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("AB".getBytes()));"#);
    assert_eq!(out, vec!["QUI="]);
}

#[test]
fn base64_encode_three_bytes_abc() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("ABC".getBytes()));"#);
    assert_eq!(out, vec!["QUJD"]);
}

#[test]
fn base64_decode_empty_string() {
    let out = run_main(r#"System.out.println(java.util.Base64.getDecoder().decode("").length);"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn base64_encode_byte_array_length() {
    let out = run_main(r#"byte[] out = java.util.Base64.getEncoder().encode("x".getBytes()); System.out.println(out.length);"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn base64_get_encoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_get_decoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getDecoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_get_url_encoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getUrlEncoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_get_url_decoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getUrlDecoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_get_mime_encoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getMimeEncoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_get_mime_decoder_not_null() {
    let out = run_main(r#"System.out.println(java.util.Base64.getMimeDecoder() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_encode_numbers() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("12345".getBytes()));"#);
    assert_eq!(out, vec!["MTIzNDU="]);
}

#[test]
fn base64_decode_numbers() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getDecoder().decode("MTIzNDU=")));"#);
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn base64_encode_unicode_ascii() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("caf\u00e9".getBytes(java.nio.charset.StandardCharsets.ISO_8859_1)));"#);
    assert_eq!(out, vec!["Y2Fmw6k="]);
}

#[test]
fn base64_url_encode_plus_replaced() {
    let out = run_main(r#"byte[] enc = java.util.Base64.getUrlEncoder().encode(">>>".getBytes()); System.out.println(new String(enc).contains("+"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn base64_url_encode_slash_replaced() {
    let out = run_main(r#"byte[] enc = java.util.Base64.getUrlEncoder().encode(new byte[]{(byte)0xff, (byte)0xfe}); System.out.println(new String(enc).contains("/"));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn base64_encode_all_zeros() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString(new byte[]{0, 0, 0}));"#);
    assert_eq!(out, vec!["AAAAAA=="]);
}

#[test]
fn base64_decode_all_zeros() {
    let out = run_main(r#"byte[] dec = java.util.Base64.getDecoder().decode("AAAAAA=="); System.out.println(dec[0] + dec[1] + dec[2]);"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn base64_encode_max_byte() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString(new byte[]{(byte)255}));"#);
    assert_eq!(out, vec!["/w=="]);
}

#[test]
fn base64_mime_encoder_line_length() {
    let out = run_main(r#"System.out.println(java.util.Base64.getMimeEncoder(4, new byte[]{'\n'}).encodeToString("abcdefghij".getBytes()).contains("\n"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_mime_decoder_ignores_whitespace() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getMimeDecoder().decode("SGVs\nbG8=")));"#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn base64_encode_java_word() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("Java".getBytes()));"#);
    assert_eq!(out, vec!["SmF2YQ=="]);
}

#[test]
fn base64_decode_java_word() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getDecoder().decode("SmF2YQ==")));"#);
    assert_eq!(out, vec!["Java"]);
}

#[test]
fn base64_encode_equals_sign() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("=".getBytes()));"#);
    assert_eq!(out, vec!["PQ=="]);
}

#[test]
fn base64_encode_newline_char() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("\n".getBytes()));"#);
    assert_eq!(out, vec!["Cg=="]);
}

#[test]
fn base64_encode_tab_char() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("\t".getBytes()));"#);
    assert_eq!(out, vec!["HQ=="]);
}

#[test]
fn base64_encode_space_char() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString(" ".getBytes()));"#);
    assert_eq!(out, vec!["IA=="]);
}

#[test]
fn base64_without_padding_shorter() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().withoutPadding().encodeToString("A".getBytes()).length());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn base64_encode_four_bytes() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("test".getBytes()));"#);
    assert_eq!(out, vec!["dGVzdA=="]);
}

#[test]
fn base64_decode_four_bytes() {
    let out = run_main(r#"System.out.println(new String(java.util.Base64.getDecoder().decode("dGVzdA==")));"#);
    assert_eq!(out, vec!["test"]);
}

#[test]
fn base64_url_without_padding_decode() {
    let out = run_main(r#"String enc = java.util.Base64.getUrlEncoder().withoutPadding().encodeToString("go".getBytes()); System.out.println(new String(java.util.Base64.getUrlDecoder().decode(enc)));"#);
    assert_eq!(out, vec!["go"]);
}

#[test]
fn base64_encode_binary_pattern() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString(new byte[]{1, 2, 3, 4}));"#);
    assert_eq!(out, vec!["AQIDBA=="]);
}

#[test]
fn base64_decode_binary_pattern() {
    let out = run_main(r#"byte[] d = java.util.Base64.getDecoder().decode("AQIDBA=="); System.out.println(d[2]);"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn base64_mime_encoder_default_rfc() {
    let out = run_main(r#"System.out.println(java.util.Base64.getMimeEncoder().encodeToString("x".getBytes()).length() >= 4);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn base64_roundtrip_empty_via_url() {
    let out = run_main(r#"String e = java.util.Base64.getUrlEncoder().encodeToString(new byte[0]); System.out.println(java.util.Base64.getUrlDecoder().decode(e).length);"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn base64_encode_slash_char() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("/".getBytes()));"#);
    assert_eq!(out, vec!["Lw=="]);
}

#[test]
fn base64_encode_plus_char() {
    let out = run_main(r#"System.out.println(java.util.Base64.getEncoder().encodeToString("+".getBytes()));"#);
    assert_eq!(out, vec!["Kw=="]);
}

