//! log package (Print, Println, Fatal) and flag package (String, Int, Bool, Parse).
//!
//! Distinct from fmt I/O in other test files; focuses on stdlib logging and CLI flags.

go_run_cases! {
    log_print_single_string => (
        "package main; import \"fmt\"; import \"log\"; func main() { log.Print(\"alpha\"); fmt.Println(\"done\") }",
        vec!["alpha", "done"]
    ),
    log_print_two_strings_no_newline => (
        "package main; import \"fmt\"; import \"log\"; func main() { log.Print(\"a\", \"b\"); fmt.Println(\"end\") }",
        vec!["ab", "end"]
    ),
    log_println_single_string => (
        "package main; import \"fmt\"; import \"log\"; func main() { log.Println(\"line\"); fmt.Println(\"ok\") }",
        vec!["line", "ok"]
    ),
    log_println_mixed_int_and_string => (
        "package main; import \"fmt\"; import \"log\"; func main() { log.Println(\"n=\", 42); fmt.Println(\"tail\") }",
        vec!["n= 42", "tail"]
    ),
    log_printf_formatted_output => (
        "package main; import \"fmt\"; import \"log\"; func main() { log.Printf(\"%s-%d\", \"vybe\", 7); fmt.Println(\"after\") }",
        vec!["vybe-7", "after"]
    ),
    flag_string_default_before_parse => (
        "package main; import \"fmt\"; import \"flag\"; func main() { name := flag.String(\"name\", \"guest\", \"user name\"); fmt.Println(*name) }",
        vec!["guest"]
    ),
    flag_int_default_before_parse => (
        "package main; import \"fmt\"; import \"flag\"; func main() { port := flag.Int(\"port\", 8080, \"listen port\"); fmt.Println(*port) }",
        vec!["8080"]
    ),
    flag_bool_default_true_before_parse => (
        "package main; import \"fmt\"; import \"flag\"; func main() { verbose := flag.Bool(\"verbose\", true, \"verbose mode\"); fmt.Println(*verbose) }",
        vec!["true"]
    ),
    flag_bool_default_false_before_parse => (
        "package main; import \"fmt\"; import \"flag\"; func main() { debug := flag.Bool(\"debug\", false, \"debug mode\"); fmt.Println(*debug) }",
        vec!["false"]
    ),
    flag_multiple_defaults_independent => (
        "package main; import \"fmt\"; import \"flag\"; func main() { host := flag.String(\"host\", \"localhost\", \"\"); port := flag.Int(\"port\", 3000, \"\"); fmt.Println(*host); fmt.Println(*port) }",
        vec!["localhost", "3000"]
    ),
    flag_string_var_rebinds_pointer => (
        "package main; import \"fmt\"; import \"flag\"; func main() { mode := flag.String(\"mode\", \"dev\", \"\"); *mode = \"prod\"; fmt.Println(*mode) }",
        vec!["prod"]
    ),
    flag_int_var_rebinds_pointer => (
        "package main; import \"fmt\"; import \"flag\"; func main() { level := flag.Int(\"level\", 1, \"\"); *level = 9; fmt.Println(*level) }",
        vec!["9"]
    ),
}

go_compile_cases! {
    log_fatal_single_message_compile =>
        "package main; import \"log\"; func main() { log.Fatal(\"abort\") }",
    log_fatal_with_code_compile =>
        "package main; import \"log\"; func main() { log.Fatal(\"exit\", 1) }",
    log_fatalf_formatted_compile =>
        "package main; import \"log\"; func main() { log.Fatalf(\"code %d\", 9) }",
    log_set_prefix_before_print_compile =>
        "package main; import \"log\"; func main() { log.SetPrefix(\"[app] \"); log.Print(\"ready\") }",
    log_set_flags_discard_date_compile =>
        "package main; import \"log\"; func main() { log.SetFlags(0); log.Print(\"plain\") }",
    log_flags_date_and_time_compile =>
        "package main; import \"log\"; func main() { log.SetFlags(log.Ldate | log.Ltime); log.Print(\"stamp\") }",
    log_output_redirect_writer_compile =>
        "package main; import \"log\"; import \"os\"; func main() { log.SetOutput(os.Stderr); log.Print(\"err\") }",
    log_print_inside_defer_compile =>
        "package main; import \"log\"; func main() { defer log.Print(\"bye\"); log.Print(\"hi\") }",

    flag_parse_no_args_compile =>
        "package main; import \"flag\"; func main() { flag.Parse() }",
    flag_parse_after_three_definitions_compile =>
        "package main; import \"flag\"; func main() { _ = flag.String(\"host\", \"\", \"\"); _ = flag.Int(\"port\", 0, \"\"); _ = flag.Bool(\"tls\", false, \"\"); flag.Parse() }",
    flag_string_var_package_scope_compile =>
        "package main; import \"flag\"; var region = flag.String(\"region\", \"us\", \"region code\"); func main() { flag.Parse(); _ = *region }",
    flag_int_var_package_scope_compile =>
        "package main; import \"flag\"; var workers = flag.Int(\"workers\", 4, \"worker count\"); func main() { flag.Parse(); _ = *workers }",
    flag_bool_var_package_scope_compile =>
        "package main; import \"flag\"; var dryRun = flag.Bool(\"dry-run\", false, \"skip writes\"); func main() { flag.Parse(); _ = *dryRun }",
    flag_parse_in_init_compile =>
        "package main; import \"flag\"; func init() { flag.Parse() }; func main() {}",
    flag_lookup_after_parse_compile =>
        "package main; import \"flag\"; func main() { flag.String(\"color\", \"red\", \"\"); flag.Parse(); _ = flag.Lookup(\"color\") }",
    flag_narg_after_parse_compile =>
        "package main; import \"flag\"; func main() { flag.Parse(); _ = flag.NArg() }",
    flag_nflag_after_parse_compile =>
        "package main; import \"flag\"; func main() { _ = flag.Bool(\"v\", false, \"\"); flag.Parse(); _ = flag.NFlag() }",
    flag_args_after_parse_compile =>
        "package main; import \"flag\"; func main() { flag.Parse(); _ = flag.Args() }",
    flag_set_with_string_value_compile =>
        "package main; import \"flag\"; func main() { f := flag.String(\"mode\", \"dev\", \"\"); _ = flag.Set(\"mode\", \"prod\"); _ = *f }",
}
