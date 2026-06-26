//! mime, html/template, runtime, image — one smoke per distinct API.

use crate::helpers::*;

go_compile_cases! {
    mime_type_by_extension => "package main; import \"mime\"; func main() { _ = mime.TypeByExtension(\".html\") }",
    mime_format_media_type => "package main; import \"mime\"; func main() { _ = mime.FormatMediaType(\"text/html\", map[string]string{\"charset\": \"utf-8\"}) }",
    multipart_new_writer => "package main; import \"mime/multipart\"; import \"bytes\"; func main() { _ = multipart.NewWriter(bytes.NewBuffer(nil)) }",
    html_template_parse => "package main; import \"html/template\"; func main() { _, _ = html/template.New(\"p\").Parse(\"{{.}}\") }",
    html_template_escape => "package main; import \"html/template\"; func main() { _ = html/template.HTMLEscapeString(\"<b>\") }",
    runtime_caller => "package main; import \"runtime\"; func main() { _, _, _, _ = runtime.Caller(0) }",
    runtime_gomaxprocs => "package main; import \"runtime\"; func main() { _ = runtime.GOMAXPROCS(0) }",
    runtime_version => "package main; import \"runtime\"; func main() { _ = runtime.Version() }",
    debug_read_build_info => "package main; import \"runtime/debug\"; func main() { _, _ = debug.ReadBuildInfo() }",
    image_new_rgba => "package main; import \"image\"; func main() { _ = image.NewRGBA(image.Rect(0, 0, 2, 2)) }",
    color_rgba_model => "package main; import \"image/color\"; func main() { _ = color.RGBAModel.Convert(color.RGBA{R: 255}) }",
    png_encode => "package main; import \"image/png\"; import \"image\"; import \"bytes\"; func main() { _ = png.Encode(bytes.NewBuffer(nil), image.NewRGBA(image.Rect(0, 0, 1, 1))) }",
}
