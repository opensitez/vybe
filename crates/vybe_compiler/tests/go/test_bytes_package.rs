//! bytes package: Compare, Contains, Index, Trim, ToUpper, Join.

go_run_cases! {
    bytes_compare_equal => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare([]byte(\"a\"), []byte(\"a\"))) }", vec!["0"]),
    bytes_compare_less => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare([]byte(\"a\"), []byte(\"b\"))) }", vec!["-1"]),
    bytes_compare_greater => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare([]byte(\"z\"), []byte(\"a\"))) }", vec!["1"]),
    bytes_contains_subslice => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Contains([]byte(\"hello\"), []byte(\"ell\"))) }", vec!["true"]),
    bytes_contains_missing => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Contains([]byte(\"go\"), []byte(\"rust\"))) }", vec!["false"]),
    bytes_index_found => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte(\"abc\"), []byte(\"b\"))) }", vec!["1"]),
    bytes_index_missing => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte(\"abc\"), []byte(\"z\"))) }", vec!["-1"]),
    bytes_trim_space => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.TrimSpace([]byte(\"  x  \")))) }", vec!["x"]),
    bytes_to_upper => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.ToUpper([]byte(\"ab\")))) }", vec!["AB"]),
    bytes_join => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.Join([][]byte{[]byte(\"a\"), []byte(\"b\")}, []byte(\"-\")))) }", vec!["a-b"]),
    bytes_equal => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Equal([]byte(\"x\"), []byte(\"x\"))) }", vec!["true"]),
    bytes_has_prefix => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.HasPrefix([]byte(\"golang\"), []byte(\"go\"))) }", vec!["true"]),
    bytes_has_suffix => ("package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.HasSuffix([]byte(\"golang\"), []byte(\"lang\"))) }", vec!["true"]),
}

go_compile_cases! {
    bytes_replace_all => "package main; import \"bytes\"; func main() { _ = bytes.ReplaceAll([]byte(\"a.a\"), []byte(\".\"), []byte(\"-\")) }",
    bytes_split_n => "package main; import \"bytes\"; func main() { _ = bytes.SplitN([]byte(\"a,b,c\"), []byte(\",\"), 2) }",
    bytes_fields => "package main; import \"bytes\"; func main() { _ = bytes.Fields([]byte(\"a  b\")) }",
    bytes_map => "package main; import \"bytes\"; func main() { _ = bytes.Map(func(r rune) rune { return r }, []byte(\"abc\")) }",
    bytes_repeat => "package main; import \"bytes\"; func main() { _ = bytes.Repeat([]byte(\"go\"), 2) }",
    bytes_trim_prefix => "package main; import \"bytes\"; func main() { _ = bytes.TrimPrefix([]byte(\"go\"), []byte(\"g\")) }",
    bytes_trim_suffix => "package main; import \"bytes\"; func main() { _ = bytes.TrimSuffix([]byte(\"go\"), []byte(\"o\")) }",
}
