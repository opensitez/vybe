//! os, path/filepath, and os/exec: environment lookup, argv, path helpers, subprocess spawn.

use crate::helpers::*;

go_run_cases! {
    os_args_len_is_zero => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(len(os.Args)) }",
        vec!["0"]
    ),
    os_args_empty_guard_prints_empty => (
        "package main; import \"fmt\"; import \"os\"; func main() { if len(os.Args) > 0 { fmt.Println(os.Args[0]) } else { fmt.Println(\"empty\") } }",
        vec!["empty"]
    ),
    os_args_join_empty_slice => (
        "package main; import \"fmt\"; import \"os\"; import \"strings\"; func main() { fmt.Println(strings.Join(os.Args, \",\")) }",
        vec![""]
    ),
}

go_compile_cases! {
    // os.Getenv
    getenv_lookup_assign => "package main; import \"os\"; func main() { v := os.Getenv(\"HOME\"); _ = v }",
    getenv_compare_empty_string => "package main; import \"os\"; func main() { _ = os.Getenv(\"MISSING\") == \"\" }",
    getenv_in_boolean_condition => "package main; import \"os\"; func main() { if os.Getenv(\"DEBUG\") != \"\" { _ = 1 } }",
    getenv_concatenated_with_literal => "package main; import \"os\"; func main() { _ = os.Getenv(\"PREFIX\") + \"_suffix\" }",
    getenv_passed_to_local_helper => "package main; import \"os\"; func pick(key string) string { return os.Getenv(key) }; func main() { _ = pick(\"PATH\") }",
    getenv_after_setenv_call => "package main; import \"os\"; func main() { os.Setenv(\"VYBE_TEST\", \"1\"); _ = os.Getenv(\"VYBE_TEST\") }",

    // os.Args
    os_args_index_first_element => "package main; import \"os\"; func main() { _ = os.Args[0] }",
    os_args_len_used_in_make => "package main; import \"os\"; func main() { buf := make([]string, len(os.Args)); _ = buf }",
    os_args_range_iteration => "package main; import \"os\"; func main() { for _, arg := range os.Args { _ = arg } }",
    os_args_copy_into_slice => "package main; import \"os\"; func main() { copied := make([]string, len(os.Args)); copy(copied, os.Args) }",
    os_args_join_with_separator => "package main; import \"os\"; import \"strings\"; func main() { _ = strings.Join(os.Args, \" \") }",
    os_args_append_spread_into_local => "package main; import \"os\"; func main() { local := append([]string{\"prog\"}, os.Args...) }",

    // path/filepath
    filepath_join_two_segments => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(\"dir\", \"file.txt\") }",
    filepath_join_three_variadic => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(\"a\", \"b\", \"c\") }",
    filepath_join_nested_calls => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(filepath.Join(\"root\", \"sub\"), \"leaf\") }",
    filepath_join_dot_relative_segment => "package main; import \"path/filepath\"; func main() { _ = filepath.Join(\".\", \"config\", \"app.toml\") }",
    filepath_base_strips_directory => "package main; import \"path/filepath\"; func main() { _ = filepath.Base(\"/var/log/app.log\") }",
    filepath_base_trailing_separator => "package main; import \"path/filepath\"; func main() { _ = filepath.Base(\"/tmp/build/\") }",
    filepath_dir_of_dotted_name => "package main; import \"path/filepath\"; func main() { _ = filepath.Dir(\"/opt/bin/tool.exe\") }",
    filepath_ext_returns_suffix => "package main; import \"path/filepath\"; func main() { _ = filepath.Ext(\"archive.tar.gz\") }",
    filepath_ext_no_extension => "package main; import \"path/filepath\"; func main() { _ = filepath.Ext(\"README\") }",
    filepath_ext_of_joined_path => "package main; import \"path/filepath\"; func main() { _ = filepath.Ext(filepath.Join(\"src\", \"main.go\")) }",

    // os/exec
    exec_command_no_extra_args => "package main; import \"os/exec\"; func main() { _ = exec.Command(\"true\") }",
    exec_command_with_multiple_args => "package main; import \"os/exec\"; func main() { _ = exec.Command(\"sh\", \"-c\", \"echo hi\") }",
    exec_command_variadic_slice_spread => "package main; import \"os/exec\"; func main() { flags := []string{\"-n\"}; _ = exec.Command(\"wc\", flags...) }",
    exec_command_name_from_filepath_base => "package main; import \"os/exec\"; import \"path/filepath\"; func main() { _ = exec.Command(filepath.Base(\"/bin/echo\"), \"vybe\") }",
    exec_command_stored_in_variable => "package main; import \"os/exec\"; func main() { cmd := exec.Command(\"date\"); _ = cmd }",
    exec_command_path_built_with_join => "package main; import \"os/exec\"; import \"path/filepath\"; func main() { bin := filepath.Join(\"usr\", \"bin\", \"env\"); _ = exec.Command(bin, \"sh\") }",
}
