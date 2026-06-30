//! os process environment: Getpid, Args, Environ, Getenv helpers, Stat/Lstat, MkdirTemp, UserHomeDir.

use crate::helpers::*;

go_run_cases! {
    os_getpid_positive => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(os.Getpid() > 0) }",
        vec!["true"]
    ),
    os_getppid_non_negative => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(os.Getppid() >= 0) }",
        vec!["true"]
    ),
    os_args_len_at_least_zero => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(len(os.Args) >= 0) }",
        vec!["true"]
    ),
    os_getenv_missing_returns_empty => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(os.Getenv(\"VYBE_NONEXISTENT_VAR_XYZ\") == \"\") }",
        vec!["true"]
    ),
    os_getenv_with_default_helper => (
        "package main; import \"fmt\"; import \"os\"; func getenvDefault(key, def string) string { if v := os.Getenv(key); v != \"\" { return v }; return def }; func main() { fmt.Println(getenvDefault(\"VYBE_MISSING_KEY_ABC\", \"fallback\")) }",
        vec!["fallback"]
    ),
    os_getenv_with_default_returns_set_value => (
        "package main; import \"fmt\"; import \"os\"; func getenvDefault(key, def string) string { if v := os.Getenv(key); v != \"\" { return v }; return def }; func main() { os.Setenv(\"VYBE_TMP_TEST_KEY\", \"live\"); fmt.Println(getenvDefault(\"VYBE_TMP_TEST_KEY\", \"fallback\")) }",
        vec!["live"]
    ),
    os_stat_current_dir_name => (
        "package main; import \"fmt\"; import \"os\"; func main() { fi, err := os.Stat(\".\"); if err != nil { fmt.Println(\"err\"); return }; fmt.Println(fi.Name() != \"\") }",
        vec!["true"]
    ),
    os_stat_current_dir_is_dir => (
        "package main; import \"fmt\"; import \"os\"; func main() { fi, err := os.Stat(\".\"); if err != nil { fmt.Println(false); return }; fmt.Println(fi.IsDir()) }",
        vec!["true"]
    ),
    os_lstat_current_dir_succeeds => (
        "package main; import \"fmt\"; import \"os\"; func main() { _, err := os.Lstat(\".\"); fmt.Println(err == nil) }",
        vec!["true"]
    ),
    os_environ_slice_non_nil => (
        "package main; import \"fmt\"; import \"os\"; func main() { fmt.Println(len(os.Environ()) >= 0) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    os_getpid_assign_int => "package main; import \"os\"; func main() { pid := os.Getpid(); _ = pid }",
    os_getppid_assign_int => "package main; import \"os\"; func main() { ppid := os.Getppid(); _ = ppid }",
    os_getuid_compile => "package main; import \"os\"; func main() { _ = os.Getuid() }",
    os_getgid_compile => "package main; import \"os\"; func main() { _ = os.Getgid() }",
    os_geteuid_compile => "package main; import \"os\"; func main() { _ = os.Geteuid() }",
    os_getegid_compile => "package main; import \"os\"; func main() { _ = os.Getegid() }",
    os_args_first_element_string => "package main; import \"os\"; func main() { if len(os.Args) > 0 { _ = os.Args[0] } }",
    os_args_index_bounds_check => "package main; import \"os\"; func main() { for i := range os.Args { _ = os.Args[i] } }",
    os_args_copy_to_local_slice => "package main; import \"os\"; func main() { dup := append([]string(nil), os.Args...); _ = dup }",
    os_environ_range_loop => "package main; import \"os\"; func main() { for _, entry := range os.Environ() { _ = entry } }",
    os_environ_lookup_prefix => "package main; import \"os\"; import \"strings\"; func main() { for _, e := range os.Environ() { if strings.HasPrefix(e, \"PATH=\") { _ = e } } }",
    os_getenv_in_string_concat => "package main; import \"os\"; func main() { _ = \"prefix_\" + os.Getenv(\"HOME\") }",
    os_getenv_passed_to_helper => "package main; import \"os\"; func lookup(k string) string { return os.Getenv(k) }; func main() { _ = lookup(\"USER\") }",
    os_getenv_default_via_or_pattern => "package main; import \"os\"; func main() { v := os.Getenv(\"VYBE_OR_DEFAULT\"); if v == \"\" { v = \"default\" }; _ = v }",
    os_setenv_then_getenv => "package main; import \"os\"; func main() { os.Setenv(\"VYBE_SET_TEST\", \"42\"); _ = os.Getenv(\"VYBE_SET_TEST\") }",
    os_unsetenv_after_set => "package main; import \"os\"; func main() { os.Setenv(\"VYBE_UNSET_TEST\", \"1\"); os.Unsetenv(\"VYBE_UNSET_TEST\"); _ = os.Getenv(\"VYBE_UNSET_TEST\") }",
    os_clearenv_compile => "package main; import \"os\"; func main() { os.Clearenv(); _ = len(os.Environ()) }",
    os_expand_env_with_dollar => "package main; import \"os\"; func main() { _ = os.ExpandEnv(\"${HOME}/bin\") }",
    os_expand_with_mapping_func => "package main; import \"os\"; func main() { _ = os.Expand(\"$USER\", func(k string) string { return os.Getenv(k) }) }",
    os_stat_file_mode_bits => "package main; import \"os\"; func main() { fi, err := os.Stat(\".\"); if err == nil { _ = fi.Mode() } }",
    os_stat_mod_time_field => "package main; import \"os\"; func main() { fi, err := os.Stat(\".\"); if err == nil { _ = fi.ModTime() } }",
    os_stat_size_field => "package main; import \"os\"; func main() { fi, err := os.Stat(\".\"); if err == nil { _ = fi.Size() } }",
    os_lstat_symlink_target => "package main; import \"os\"; func main() { _, _ = os.Lstat(\".\") }",
    os_stat_is_not_exist_check => "package main; import \"os\"; func main() { _, err := os.Stat(\"/no/such/vybe/path\"); _ = os.IsNotExist(err) }",
    os_stat_is_permission_check => "package main; import \"os\"; func main() { _, err := os.Stat(\"/etc/shadow\"); _ = os.IsPermission(err) }",
    os_mkdir_temp_with_pattern => "package main; import \"os\"; func main() { dir, err := os.MkdirTemp(\"\", \"vybe-test-*\"); if err == nil { defer os.RemoveAll(dir) }; _ = dir }",
    os_mkdir_temp_in_current_dir => "package main; import \"os\"; func main() { dir, err := os.MkdirTemp(\".\", \"local-*\"); if err == nil { defer os.RemoveAll(dir) }; _ = dir }",
    os_mkdir_all_nested_path => "package main; import \"os\"; func main() { _ = os.MkdirAll(\"/tmp/vybe/nested/dir\", 0755) }",
    os_user_home_dir_compile => "package main; import \"os\"; func main() { home, err := os.UserHomeDir(); _ = home; _ = err }",
    os_user_cache_dir_compile => "package main; import \"os\"; func main() { _, _ = os.UserCacheDir() }",
    os_user_config_dir_compile => "package main; import \"os\"; func main() { _, _ = os.UserConfigDir() }",
    os_temp_dir_compile => "package main; import \"os\"; func main() { _ = os.TempDir() }",
    os_getwd_current_working => "package main; import \"os\"; func main() { _, _ = os.Getwd() }",
    os_chdir_then_getwd => "package main; import \"os\"; func main() { wd, _ := os.Getwd(); defer os.Chdir(wd); _ = os.Chdir(\".\") }",
    os_hostname_compile => "package main; import \"os\"; func main() { _, _ = os.Hostname() }",
    os_executable_compile => "package main; import \"os\"; func main() { _, _ = os.Executable() }",
    os_read_file_compile => "package main; import \"os\"; func main() { _, _ = os.ReadFile(\"/etc/hosts\") }",
    os_write_file_compile => "package main; import \"os\"; func main() { _ = os.WriteFile(\"/tmp/vybe-write-test.txt\", []byte(\"x\"), 0644) }",
    os_remove_file_compile => "package main; import \"os\"; func main() { _ = os.Remove(\"/tmp/vybe-write-test.txt\") }",
    os_rename_file_compile => "package main; import \"os\"; func main() { _ = os.Rename(\"/tmp/a\", \"/tmp/b\") }",
    os_create_temp_file => "package main; import \"os\"; func main() { f, err := os.CreateTemp(\"\", \"vybe-*\"); if err == nil { defer os.Remove(f.Name()); defer f.Close() } }",
    os_open_file_readonly => "package main; import \"os\"; func main() { f, err := os.Open(\".\"); if err == nil { defer f.Close() } }",
    os_open_file_with_flags => "package main; import \"os\"; func main() { _, _ = os.OpenFile(\"/tmp/vybe-open\", os.O_CREATE|os.O_RDWR, 0644) }",
    os_pipe_create => "package main; import \"os\"; func main() { r, w, err := os.Pipe(); if err == nil { r.Close(); w.Close() } }",
    os_process_state_exited => "package main; import \"os\"; import \"os/exec\"; func main() { cmd := exec.Command(\"true\"); err := cmd.Run(); if err == nil { _ = cmd.ProcessState.Exited() } }",
    os_same_file_stat_compare => "package main; import \"os\"; func main() { a, e1 := os.Stat(\".\"); b, e2 := os.Lstat(\".\"); if e1 == nil && e2 == nil { _, _ = os.SameFile(a, b) } }",
}
