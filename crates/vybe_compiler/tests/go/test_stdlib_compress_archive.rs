//! compress/* and archive/* — one compile smoke per distinct API.


go_compile_cases! {
    flate_new_writer => "package main; import \"compress/flate\"; import \"bytes\"; func main() { _, _ = flate.NewWriter(bytes.NewBuffer(nil), 6) }",
    flate_new_reader => "package main; import \"compress/flate\"; import \"bytes\"; func main() { _ = flate.NewReader(bytes.NewReader(nil)) }",
    gzip_new_writer => "package main; import \"compress/gzip\"; import \"bytes\"; func main() { _ = gzip.NewWriter(bytes.NewBuffer(nil)) }",
    gzip_new_reader => "package main; import \"compress/gzip\"; import \"bytes\"; func main() { _, _ = gzip.NewReader(bytes.NewReader(nil)) }",
    zlib_new_writer => "package main; import \"compress/zlib\"; import \"bytes\"; func main() { _ = zlib.NewWriter(bytes.NewBuffer(nil)) }",
    zlib_new_reader => "package main; import \"compress/zlib\"; import \"bytes\"; func main() { _ = zlib.NewReader(bytes.NewReader(nil)) }",
    tar_new_writer => "package main; import \"archive/tar\"; import \"bytes\"; func main() { _ = tar.NewWriter(bytes.NewBuffer(nil)) }",
    tar_new_reader => "package main; import \"archive/tar\"; import \"bytes\"; func main() { _ = tar.NewReader(bytes.NewReader(nil)) }",
    zip_new_writer => "package main; import \"archive/zip\"; import \"bytes\"; func main() { _ = zip.NewWriter(bytes.NewBuffer(nil)) }",
    zip_new_reader => "package main; import \"archive/zip\"; import \"bytes\"; func main() { _, _ = zip.NewReader(bytes.NewReader(nil), 0) }",
}
