//! embed and unsafe.Sizeof/Alignof compile patterns.

go_compile_cases! {
    unsafe_sizeof_int => "package main; import \"unsafe\"; func main() { _ = unsafe.Sizeof(int(0)) }",
    unsafe_sizeof_struct => "package main; import \"unsafe\"; type S struct { a int; b byte }; func main() { _ = unsafe.Sizeof(S{}) }",
    unsafe_alignof_int => "package main; import \"unsafe\"; func main() { _ = unsafe.Alignof(int(0)) }",
    unsafe_offsetof_field => "package main; import \"unsafe\"; type S struct { a int; b int }; func main() { _ = unsafe.Offsetof(S{}.b) }",
    unsafe_pointer_convert => "package main; import \"unsafe\"; func main() { var x int; p := unsafe.Pointer(&x); _ = p }",
    embed_string_var => "package main; import _ \"embed\"; //go:embed hello\nvar s string\nfunc main() { _ = s }",
    embed_bytes_var => "package main; import _ \"embed\"; //go:embed data.bin\nvar b []byte\nfunc main() { _ = b }",
    embed_fs_var => "package main; import \"embed\"; import \"io/fs\"; //go:embed templates/*\nvar tmpl embed.FS\nfunc main() { _, _ = fs.ReadFile(tmpl, \"templates/x\") }",
}
