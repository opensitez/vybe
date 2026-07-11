//! io/ioutil and io/fs compile coverage (legacy ioutil + fs interfaces).

go_compile_cases! {
    ioutil_read_all => "package main; import \"io/ioutil\"; import \"strings\"; func main() { _, _ = ioutil.ReadAll(strings.NewReader(\"hi\")) }",
    ioutil_read_file_compile => "package main; import \"io/ioutil\"; func main() { _, _ = ioutil.ReadFile(\"/dev/null\") }",
    ioutil_write_file_compile => "package main; import \"io/ioutil\"; func main() { _ = ioutil.WriteFile(\"out.txt\", []byte(\"x\"), 0644) }",
    ioutil_nop_closer => "package main; import \"io/ioutil\"; func main() { _ = ioutil.NopCloser(nil) }",
    io_copy_compile => "package main; import \"io\"; import \"strings\"; func main() { _, _ = io.Copy(io.Discard, strings.NewReader(\"a\")) }",
    io_read_all => "package main; import \"io\"; import \"strings\"; func main() { _, _ = io.ReadAll(strings.NewReader(\"ab\")) }",
    io_pipe_compile => "package main; import \"io\"; func main() { r, w := io.Pipe(); _ = r; _ = w }",
    io_limit_reader => "package main; import \"io\"; import \"strings\"; func main() { _ = io.LimitReader(strings.NewReader(\"abcd\"), 2) }",
    io_multi_writer => "package main; import \"io\"; import \"os\"; func main() { _ = io.MultiWriter(os.Stdout, os.Stderr) }",
    io_tee_reader => "package main; import \"io\"; import \"os\"; import \"strings\"; func main() { _ = io.TeeReader(strings.NewReader(\"z\"), os.Stdout) }",
    fs_walk_dir_compile => "package main; import \"io/fs\"; import \"os\"; func main() { _ = fs.WalkDir(os.DirFS(\".\"), \".\", func(path string, d fs.DirEntry, err error) error { return nil }) }",
    fs_read_file_compile => "package main; import \"io/fs\"; import \"os\"; func main() { _, _ = fs.ReadFile(os.DirFS(\".\"), \"go.mod\") }",
    fs_glob_compile => "package main; import \"io/fs\"; import \"os\"; func main() { _, _ = fs.Glob(os.DirFS(\".\"), \"*.go\") }",
    fs_sub_compile => "package main; import \"io/fs\"; import \"os\"; func main() { _, _ = fs.Sub(os.DirFS(\".\"), \"crates\") }",
}

go_run_cases! {
    io_read_byte => ("package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { r := strings.NewReader(\"A\"); b, err := r.ReadByte(); fmt.Println(string(b), err == nil) }", vec!["A true"]),
    io_write_string => ("package main; import \"fmt\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; n, _ := io.WriteString(&buf, \"go\"); fmt.Println(n, buf.String()) }", vec!["2 go"]),
}
