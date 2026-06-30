//! mime and multipart extended: ParseMediaType, FormatMediaType, Writer CreatePart,
//! Reader NextPart, boundary parsing — distinct from `test_stdlib_mime_runtime.rs`.

use crate::helpers::*;

go_run_cases! {
    mime_parse_media_type_simple => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, params, err := mime.ParseMediaType(\"text/html\"); fmt.Println(mt); fmt.Println(err == nil); fmt.Println(len(params)) }",
        vec!["text/html", "true", "0"]
    ),
    mime_parse_media_type_with_charset => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, params, _ := mime.ParseMediaType(\"text/plain; charset=utf-8\"); fmt.Println(mt); fmt.Println(params[\"charset\"]) }",
        vec!["text/plain", "utf-8"]
    ),
    mime_parse_media_type_quoted_value => (
        "package main; import \"fmt\"; import \"mime\"; func main() { _, params, _ := mime.ParseMediaType(`multipart/form-data; boundary=\"abc123\"`); fmt.Println(params[\"boundary\"]) }",
        vec!["abc123"]
    ),
    mime_parse_media_type_multiple_params => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, params, _ := mime.ParseMediaType(\"text/html; charset=utf-8; level=1\"); fmt.Println(mt); fmt.Println(params[\"level\"]) }",
        vec!["text/html", "1"]
    ),
    mime_parse_media_type_application_json => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, _, _ := mime.ParseMediaType(\"application/json\"); fmt.Println(mt) }",
        vec!["application/json"]
    ),
    mime_format_media_type_no_params => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.FormatMediaType(\"text/plain\", nil)) }",
        vec!["text/plain"]
    ),
    mime_format_media_type_charset => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.FormatMediaType(\"text/html\", map[string]string{\"charset\": \"utf-8\"})) }",
        vec!["text/html; charset=utf-8"]
    ),
    mime_format_media_type_boundary => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.FormatMediaType(\"multipart/mixed\", map[string]string{\"boundary\": \"xyz\"})) }",
        vec!["multipart/mixed; boundary=xyz"]
    ),
    mime_format_media_type_two_params => (
        "package main; import \"fmt\"; import \"mime\"; func main() { s := mime.FormatMediaType(\"multipart/form-data\", map[string]string{\"boundary\": \"b\", \"charset\": \"utf-8\"}); fmt.Println(len(s) > 20) }",
        vec!["true"]
    ),
    mime_parse_format_roundtrip => (
        "package main; import \"fmt\"; import \"mime\"; func main() { orig := \"text/plain; charset=utf-8\"; mt, params, _ := mime.ParseMediaType(orig); back := mime.FormatMediaType(mt, params); fmt.Println(back == orig) }",
        vec!["true"]
    ),
    mime_type_by_extension_html => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.TypeByExtension(\".html\")) }",
        vec!["text/html; charset=utf-8"]
    ),
    mime_type_by_extension_json => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.TypeByExtension(\".json\")) }",
        vec!["application/json"]
    ),
    mime_extensions_by_type_plain => (
        "package main; import \"fmt\"; import \"mime\"; func main() { exts, _ := mime.ExtensionsByType(\"text/plain\"); fmt.Println(len(exts) > 0) }",
        vec!["true"]
    ),
    multipart_writer_boundary_nonempty => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { w := multipart.NewWriter(bytes.NewBuffer(nil)); fmt.Println(len(w.Boundary()) > 0) }",
        vec!["true"]
    ),
    multipart_writer_form_data_content_type => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { w := multipart.NewWriter(bytes.NewBuffer(nil)); fmt.Println(len(w.FormDataContentType()) > 0) }",
        vec!["true"]
    ),
    multipart_writer_create_part_header => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; import \"net/textproto\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); h := make(textproto.MIMEHeader); h.Set(\"Content-Type\", \"text/plain\"); p, _ := w.CreatePart(h); fmt.Println(p != nil); w.Close() }",
        vec!["true"]
    ),
    multipart_writer_create_form_field => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); p, _ := w.CreateFormField(\"name\"); fmt.Println(p != nil); w.Close() }",
        vec!["true"]
    ),
    multipart_writer_create_form_file => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); p, _ := w.CreateFormFile(\"upload\", \"file.txt\"); fmt.Println(p != nil); w.Close() }",
        vec!["true"]
    ),
    multipart_reader_next_part => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"k\"); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, err := r.NextPart(); fmt.Println(err == nil); fmt.Println(p != nil) }",
        vec!["true", "true"]
    ),
    multipart_reader_next_part_form_name => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); fw, _ := w.CreateFormField(\"token\"); fw.Write([]byte(\"abc\")); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, _ := r.NextPart(); fmt.Println(p.FormName()); b, _ := io.ReadAll(p); fmt.Println(string(b)) }",
        vec!["token", "abc"]
    ),
    multipart_reader_two_parts => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"a\"); w.CreateFormField(\"b\"); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); _, _ = r.NextPart(); p2, err := r.NextPart(); fmt.Println(err == nil); fmt.Println(p2.FormName()) }",
        vec!["true", "b"]
    ),
    multipart_reader_boundary_from_content_type => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; import \"mime\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"x\"); w.Close(); ct := w.FormDataContentType(); _, params, _ := mime.ParseMediaType(ct); r := multipart.NewReader(&buf, params[\"boundary\"]); p, err := r.NextPart(); fmt.Println(err == nil); fmt.Println(p != nil) }",
        vec!["true", "true"]
    ),
    multipart_reader_is_boundary_error => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); _, _ = r.NextPart(); _, err := r.NextPart(); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    mime_parse_media_type_image_png => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, _, _ := mime.ParseMediaType(\"image/png\"); fmt.Println(mt) }",
        vec!["image/png"]
    ),
    mime_parse_media_type_wildcard => (
        "package main; import \"fmt\"; import \"mime\"; func main() { mt, _, _ := mime.ParseMediaType(\"*/*\"); fmt.Println(mt) }",
        vec!["*/*"]
    ),
    mime_format_media_type_empty_map => (
        "package main; import \"fmt\"; import \"mime\"; func main() { fmt.Println(mime.FormatMediaType(\"application/octet-stream\", map[string]string{})) }",
        vec!["application/octet-stream"]
    ),
    multipart_writer_set_boundary => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; func main() { w := multipart.NewWriter(bytes.NewBuffer(nil)); w.SetBoundary(\"customBoundary42\"); fmt.Println(w.Boundary()) }",
        vec!["customBoundary42"]
    ),
    multipart_reader_part_content_type => (
        "package main; import \"fmt\"; import \"mime/multipart\"; import \"bytes\"; import \"net/textproto\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); h := make(textproto.MIMEHeader); h.Set(\"Content-Type\", \"application/json\"); w.CreatePart(h); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, _ := r.NextPart(); fmt.Println(p.Header.Get(\"Content-Type\")) }",
        vec!["application/json"]
    ),
}

go_compile_cases! {
    mime_parse_media_type_invalid => "package main; import \"mime\"; func main() { _, _, err := mime.ParseMediaType(\"not valid\"); _ = err }",
    mime_parse_media_type_empty => "package main; import \"mime\"; func main() { _, _, err := mime.ParseMediaType(\"\"); _ = err }",
    mime_format_media_type_special_chars => "package main; import \"mime\"; func main() { _ = mime.FormatMediaType(\"text/plain\", map[string]string{\"name\": \"file name\"}) }",
    mime_extensions_by_type_unknown => "package main; import \"mime\"; func main() { _, err := mime.ExtensionsByType(\"application/x-unknown-vybe\"); _ = err }",
    mime_type_by_extension_empty => "package main; import \"mime\"; func main() { _ = mime.TypeByExtension(\"\") }",
    mime_type_by_extension_no_dot => "package main; import \"mime\"; func main() { _ = mime.TypeByExtension(\"txt\") }",
    multipart_new_writer_nil_buffer => "package main; import \"mime/multipart\"; func main() { _ = multipart.NewWriter(nil) }",
    multipart_writer_create_part_twice => "package main; import \"mime/multipart\"; import \"bytes\"; import \"net/textproto\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); h := make(textproto.MIMEHeader); w.CreatePart(h); w.CreatePart(h); w.Close() }",
    multipart_writer_write_field_body => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); fw, _ := w.CreateFormField(\"msg\"); _, _ = fw.Write([]byte(\"hello\")); w.Close() }",
    multipart_writer_write_file_body => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); fw, _ := w.CreateFormFile(\"f\", \"a.txt\"); _, _ = fw.Write([]byte(\"data\")); w.Close() }",
    multipart_reader_next_part_eof => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); _, _ = r.NextPart(); _, _ = r.NextPart() }",
    multipart_reader_read_form => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"a\"); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); _, _ = r.ReadForm(1024) }",
    multipart_reader_part_read_all => "package main; import \"mime/multipart\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); fw, _ := w.CreateFormField(\"k\"); fw.Write([]byte(\"v\")); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, _ := r.NextPart(); _, _ = io.ReadAll(p) }",
    multipart_writer_close_flushes => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"x\"); _ = w.Close() }",
    multipart_reader_boundary_mismatch => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.Close(); r := multipart.NewReader(&buf, \"wrongBoundary\"); _, err := r.NextPart(); _ = err }",
    mime_parse_media_type_case_insensitive => "package main; import \"mime\"; func main() { mt, params, _ := mime.ParseMediaType(\"Text/HTML; Charset=UTF-8\"); _ = mt; _ = params }",
    mime_format_media_type_multipart_mixed => "package main; import \"mime\"; func main() { _ = mime.FormatMediaType(\"multipart/mixed\", map[string]string{\"boundary\": \"----WebKit\"}) }",
    multipart_writer_form_data_content_type_has_boundary => "package main; import \"mime/multipart\"; import \"bytes\"; import \"strings\"; func main() { w := multipart.NewWriter(bytes.NewBuffer(nil)); ct := w.FormDataContentType(); _ = strings.Contains(ct, \"boundary=\") }",
    multipart_reader_part_file_name => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormFile(\"doc\", \"report.pdf\"); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, _ := r.NextPart(); _ = p.FileName() }",
    mime_parse_media_type_base64_param => "package main; import \"mime\"; func main() { _, params, _ := mime.ParseMediaType(\"text/plain; charset=iso-8859-1\"); _ = params[\"charset\"] }",
    multipart_writer_multiple_files => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormFile(\"a\", \"1.txt\"); w.CreateFormFile(\"b\", \"2.txt\"); w.Close() }",
    mime_format_media_type_quoted_boundary => "package main; import \"mime\"; func main() { _ = mime.FormatMediaType(\"multipart/form-data\", map[string]string{\"boundary\": \"abc=def\"}) }",
    multipart_reader_next_part_header => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := multipart.NewWriter(&buf); w.CreateFormField(\"h\"); w.Close(); r := multipart.NewReader(&buf, w.Boundary()); p, _ := r.NextPart(); _ = p.Header }",
    mime_type_by_extension_css => "package main; import \"mime\"; func main() { _ = mime.TypeByExtension(\".css\") }",
    mime_type_by_extension_svg => "package main; import \"mime\"; func main() { _ = mime.TypeByExtension(\".svg\") }",
    multipart_writer_set_boundary_invalid => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { w := multipart.NewWriter(bytes.NewBuffer(nil)); err := w.SetBoundary(\"bad boundary spaces\"); _ = err }",
    multipart_reader_empty_body => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { r := multipart.NewReader(bytes.NewReader([]byte{}), \"b\"); _, err := r.NextPart(); _ = err }",
}
