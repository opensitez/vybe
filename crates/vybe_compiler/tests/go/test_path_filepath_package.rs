//! path and path/filepath: Join, Base, Clean, Ext, IsAbs, Split — slash vs OS paths.


go_run_cases! {
    path_join_three_forward_segments => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Join(\"a\", \"b\", \"c\")) }",
        vec!["a/b/c"]
    ),
    path_join_elides_empty_element => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Join(\"a\", \"\", \"b\")) }",
        vec!["a/b"]
    ),
    path_clean_collapses_duplicate_slashes => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Clean(\"/a//b/\")) }",
        vec!["/a/b"]
    ),
    path_clean_resolves_dot_dot_parent => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Clean(\"/a/b/../c\")) }",
        vec!["/a/c"]
    ),
    path_clean_empty_returns_dot => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Clean(\"\")) }",
        vec!["."]
    ),
    path_base_last_path_element => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Base(\"/x/y/z.txt\")) }",
        vec!["z.txt"]
    ),
    path_base_trailing_slash_strips_dir => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Base(\"/usr/\")) }",
        vec!["usr"]
    ),
    path_ext_go_source_suffix => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Ext(\"main.go\")) }",
        vec![".go"]
    ),
    path_ext_compound_name_last_dot => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.Ext(\"data.tar.gz\")) }",
        vec![".gz"]
    ),
    path_is_abs_slash_prefix_true => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.IsAbs(\"/home\")) }",
        vec!["true"]
    ),
    path_is_abs_bare_name_false => (
        "package main; import \"fmt\"; import \"path\"; func main() { fmt.Println(path.IsAbs(\"home\")) }",
        vec!["false"]
    ),
    path_split_yields_dir_and_file => (
        "package main; import \"fmt\"; import \"path\"; func main() { dir, file := path.Split(\"/a/b/c\"); fmt.Println(dir); fmt.Println(file) }",
        vec!["/a/b", "c"]
    ),
    path_split_bare_filename_no_dir => (
        "package main; import \"fmt\"; import \"path\"; func main() { dir, file := path.Split(\"file.txt\"); fmt.Println(dir); fmt.Println(file) }",
        vec![".", "file.txt"]
    ),
    filepath_clean_simplifies_dot_segments => (
        "package main; import \"fmt\"; import \"path/filepath\"; func main() { fmt.Println(filepath.Clean(\"a/b/./c/\")) }",
        vec!["a/b/c"]
    ),
    filepath_is_abs_rooted_path_true => (
        "package main; import \"fmt\"; import \"path/filepath\"; func main() { fmt.Println(filepath.IsAbs(\"/var\")) }",
        vec!["true"]
    ),
    filepath_is_abs_relative_path_false => (
        "package main; import \"fmt\"; import \"path/filepath\"; func main() { fmt.Println(filepath.IsAbs(\"local\")) }",
        vec!["false"]
    ),
    filepath_split_parent_and_leaf => (
        "package main; import \"fmt\"; import \"path/filepath\"; func main() { dir, file := filepath.Split(\"/opt/bin/go\"); fmt.Println(dir); fmt.Println(file) }",
        vec!["/opt/bin", "go"]
    ),
}

go_compile_cases! {
    path_join_double_dot_parent_segment => "package main; import \"path\"; func main() { _ = path.Join(\"/a\", \"..\", \"b\") }",
    path_clean_relative_dot_prefix => "package main; import \"path\"; func main() { _ = path.Clean(\"./a/../b\") }",
    path_ext_leading_dotfile_empty => "package main; import \"path\"; func main() { _ = path.Ext(\".profile\") }",
    path_is_abs_lone_slash_root => "package main; import \"path\"; func main() { _ = path.IsAbs(\"/\") }",
    path_split_root_slash_only => "package main; import \"path\"; func main() { _, _ = path.Split(\"/\") }",
    path_base_single_name_no_slash => "package main; import \"path\"; func main() { _ = path.Base(\"archive.zip\") }",
    filepath_join_double_dot_relative => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(\"..\", \"a\", \"b\") }",
    filepath_clean_after_join_chain => "package main; import \"path/filepath\"; func main() { _ = filepath.Clean(filepath.Join(\"a\", \"b\", \"..\", \"c\")) }",
    filepath_ext_on_cleaned_path => "package main; import \"path/filepath\"; func main() { _ = filepath.Ext(filepath.Clean(\"dir/file.GO\")) }",
    filepath_is_abs_windows_drive_letter => "package main; import \"path/filepath\"; func main() { _ = filepath.IsAbs(`C:\\Windows`) }",
    filepath_split_empty_path_tuple => "package main; import \"path/filepath\"; func main() { _, _ = filepath.Split(\"\") }",
    path_and_filepath_mixed_expression => "package main; import \"path\"; import \"path/filepath\"; func main() { _ = filepath.Clean(path.Join(\"src\", \"main.go\")) }",
    filepath_join_leading_absolute_segment => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(\"/tmp\", \"vybe\", \"out\") }",
}
