//! log: Print/Printf/Println, Fatal/Panic compile-only, SetPrefix, SetFlags,
//! Output with custom writer via bytes.Buffer — extended coverage distinct
//! from `test_log_flag_packages.rs`.

use crate::helpers::*;

go_run_cases! {
    log_print_to_custom_buffer => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"buf\"); fmt.Println(buf.String()) }",
        vec!["buf\n"]
    ),
    log_println_to_custom_buffer => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Println(\"line\"); fmt.Println(buf.String()) }",
        vec!["line\n"]
    ),
    log_printf_to_custom_buffer => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"%s-%d\", \"vybe\", 3); fmt.Println(buf.String()) }",
        vec!["vybe-3\n"]
    ),
    log_set_prefix_appears_in_output => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.SetPrefix(\"[svc] \"); log.Print(\"ping\"); fmt.Println(buf.String()) }",
        vec!["[svc] ping\n"]
    ),
    log_set_prefix_empty => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.SetPrefix(\"\"); log.Print(\"plain\"); fmt.Println(buf.String()) }",
        vec!["plain\n"]
    ),
    log_set_flags_zero_no_metadata => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"x\"); s := buf.String(); fmt.Println(len(s) == 2) }",
        vec!["true"]
    ),
    log_set_flags_date_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ldate); log.Print(\"d\"); fmt.Println(len(buf.String()) > 2) }",
        vec!["true"]
    ),
    log_set_flags_time_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ltime); log.Print(\"t\"); fmt.Println(len(buf.String()) > 2) }",
        vec!["true"]
    ),
    log_set_flags_date_or_time => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ldate | log.Ltime); log.Print(\"both\"); fmt.Println(len(buf.String()) > 4) }",
        vec!["true"]
    ),
    log_set_flags_microseconds_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ltime | log.Lmicroseconds); log.Print(\"us\"); fmt.Println(len(buf.String()) > 2) }",
        vec!["true"]
    ),
    log_set_flags_longfile_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Llongfile); log.Print(\"lf\"); fmt.Println(len(buf.String()) > 2) }",
        vec!["true"]
    ),
    log_set_flags_shortfile_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Lshortfile); log.Print(\"sf\"); fmt.Println(len(buf.String()) > 2) }",
        vec!["true"]
    ),
    log_print_multiple_args_joined => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"a\", \"b\", \"c\"); fmt.Println(buf.String()) }",
        vec!["abc\n"]
    ),
    log_println_multiple_args_spaced => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Println(\"a\", 1); fmt.Println(buf.String()) }",
        vec!["a 1\n"]
    ),
    log_printf_no_extra_newline_verb => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"n=%d\", 5); fmt.Println(buf.String()) }",
        vec!["n=5\n"]
    ),
    log_printf_bool_verb => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"ok=%t\", true); fmt.Println(buf.String()) }",
        vec!["ok=true\n"]
    ),
    log_printf_hex_verb => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"x=%x\", 255); fmt.Println(buf.String()) }",
        vec!["x=ff\n"]
    ),
    log_output_direct_call => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetFlags(0); _ = log.Output(0, \"direct\\n\"); log.SetOutput(&buf); log.SetFlags(0); log.Print(\"after\"); fmt.Println(buf.String()) }",
        vec!["after\n"]
    ),
    log_output_to_buffer_with_prefix => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.SetPrefix(\"P> \"); _ = log.Output(1, \"msg\\n\"); fmt.Println(buf.String()) }",
        vec!["P> msg\n"]
    ),
    log_print_empty_string => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"\"); fmt.Println(buf.String()) }",
        vec!["\n"]
    ),
    log_println_empty => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Println(); fmt.Println(buf.String()) }",
        vec!["\n"]
    ),
    log_printf_empty_format => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"\"); fmt.Println(buf.String()) }",
        vec!["\n"]
    ),
    log_print_int_converted => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(42); fmt.Println(buf.String()) }",
        vec!["42\n"]
    ),
    log_println_unicode => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Println(\"日\"); fmt.Println(buf.String()) }",
        vec!["日\n"]
    ),
    log_set_output_restores_default_behavior => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"captured\"); log.SetOutput(io.Discard); fmt.Println(buf.String()) }",
        vec!["captured\n"]
    ),
    log_prefix_and_message_order => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.SetPrefix(\">>\"); log.Print(\"go\"); fmt.Println(strings.HasPrefix(buf.String(), \">>\")) }",
        vec!["true"]
    ),
    log_flags_utc_bit => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ldate | log.Ltime | log.LUTC); log.Print(\"utc\"); fmt.Println(len(buf.String()) > 3) }",
        vec!["true"]
    ),
    log_flags_msgprefix_with_prefix => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetPrefix(\"pre\"); log.SetFlags(log.Lmsgprefix); log.Print(\"m\"); fmt.Println(len(buf.String()) > 1) }",
        vec!["true"]
    ),
    log_printf_quoted_string => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Printf(\"q=%q\", \"go\"); fmt.Println(buf.String()) }",
        vec!["q=\"go\"\n"]
    ),
    log_print_three_ints => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(1, 2, 3); fmt.Println(buf.String()) }",
        vec!["123\n"]
    ),
    log_println_mixed_types => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Println(\"n\", 0, true); fmt.Println(buf.String()) }",
        vec!["n 0 true\n"]
    ),
    log_set_prefix_then_clear => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.SetPrefix(\"A\"); log.SetPrefix(\"\"); log.Print(\"z\"); fmt.Println(buf.String()) }",
        vec!["z\n"]
    ),
    log_output_call_depth_two => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); _ = log.Output(2, \"depth\\n\"); fmt.Println(buf.String()) }",
        vec!["depth\n"]
    ),
    log_print_after_set_flags_date => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(log.Ldate); log.Print(\"msg\"); fmt.Println(len(buf.String()) > 3) }",
        vec!["true"]
    ),
    log_buffer_reset_between_writes => (
        "package main; import \"fmt\"; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"first\"); buf.Reset(); log.Print(\"second\"); fmt.Println(buf.String()) }",
        vec!["second\n"]
    ),
}

go_compile_cases! {
    log_fatal_single_arg => "package main; import \"log\"; func main() { log.Fatal(\"stop\") }",
    log_fatal_two_args => "package main; import \"log\"; func main() { log.Fatal(\"code\", 1) }",
    log_fatalf_format => "package main; import \"log\"; func main() { log.Fatalf(\"err=%s\", \"x\") }",
    log_fatalln_line => "package main; import \"log\"; func main() { log.Fatalln(\"bye\") }",
    log_panic_single_arg => "package main; import \"log\"; func main() { log.Panic(\"boom\") }",
    log_panic_two_args => "package main; import \"log\"; func main() { log.Panic(\"fail\", 9) }",
    log_panicf_format => "package main; import \"log\"; func main() { log.Panicf(\"p=%d\", 1) }",
    log_panicln_line => "package main; import \"log\"; func main() { log.Panicln(\"panic\") }",
    log_set_output_os_stderr => "package main; import \"log\"; import \"os\"; func main() { log.SetOutput(os.Stderr); log.Print(\"e\") }",
    log_set_output_io_discard => "package main; import \"log\"; import \"io\"; func main() { log.SetOutput(io.Discard); log.Print(\"gone\") }",
    log_set_flags_std_flags => "package main; import \"log\"; func main() { log.SetFlags(log.LstdFlags); log.Print(\"std\") }",
    log_set_flags_all_metadata => "package main; import \"log\"; func main() { log.SetFlags(log.Ldate | log.Ltime | log.Lmicroseconds | log.Llongfile); log.Print(\"all\") }",
    log_print_in_init => "package main; import \"log\"; func init() { log.Print(\"init\") }; func main() {}",
    log_println_in_defer => "package main; import \"log\"; func main() { defer log.Println(\"defer\") }",
    log_printf_in_defer => "package main; import \"log\"; func main() { defer log.Printf(\"d=%d\", 1) }",
    log_output_with_longfile_flag => "package main; import \"log\"; func main() { log.SetFlags(log.Llongfile); _ = log.Output(0, \"o\\n\") }",
    log_output_with_shortfile_flag => "package main; import \"log\"; func main() { log.SetFlags(log.Lshortfile); _ = log.Output(0, \"o\\n\") }",
    log_set_prefix_before_fatal_compile => "package main; import \"log\"; func main() { log.SetPrefix(\"X\"); log.Fatal(\"f\") }",
    log_set_prefix_before_panic_compile => "package main; import \"log\"; func main() { log.SetPrefix(\"X\"); log.Panic(\"p\") }",
    log_print_to_bytes_buffer_var => "package main; import \"log\"; import \"bytes\"; func main() { var buf bytes.Buffer; log.SetOutput(&buf); log.SetFlags(0); log.Print(\"b\") }",
}
