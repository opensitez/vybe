//! flag: String/Bool/Int/Int64/Uint/Duration vars, Parse, VisitAll, Set,
//! Lookup, CommandLine vs New FlagSet, ContinueOnError — extended coverage
//! distinct from `test_log_flag_packages.rs`.

use crate::helpers::*;

go_run_cases! {
    flag_int64_default_before_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { size := flag.Int64(\"size\", 1024, \"byte size\"); fmt.Println(*size) }",
        vec!["1024"]
    ),
    flag_int64_after_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { size := flag.Int64(\"size\", 1, \"\"); _ = flag.Set(\"size\", \"2048\"); fmt.Println(*size) }",
        vec!["2048"]
    ),
    flag_uint_default_before_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { count := flag.Uint(\"count\", 10, \"item count\"); fmt.Println(*count) }",
        vec!["10"]
    ),
    flag_uint_after_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { count := flag.Uint(\"count\", 0, \"\"); _ = flag.Set(\"count\", \"99\"); fmt.Println(*count) }",
        vec!["99"]
    ),
    flag_uint64_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { big := flag.Uint64(\"big\", 1000, \"\"); fmt.Println(*big) }",
        vec!["1000"]
    ),
    flag_uint64_after_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { big := flag.Uint64(\"big\", 0, \"\"); _ = flag.Set(\"big\", \"5000\"); fmt.Println(*big) }",
        vec!["5000"]
    ),
    flag_duration_default_zero => (
        "package main; import \"fmt\"; import \"flag\"; import \"time\"; func main() { d := flag.Duration(\"timeout\", 0, \"\"); fmt.Println(*d == 0) }",
        vec!["true"]
    ),
    flag_duration_after_set_seconds => (
        "package main; import \"fmt\"; import \"flag\"; func main() { d := flag.Duration(\"timeout\", 0, \"\"); _ = flag.Set(\"timeout\", \"2s\"); fmt.Println(*d) }",
        vec!["2s"]
    ),
    flag_duration_after_set_milliseconds => (
        "package main; import \"fmt\"; import \"flag\"; func main() { d := flag.Duration(\"wait\", 0, \"\"); _ = flag.Set(\"wait\", \"250ms\"); fmt.Println(*d) }",
        vec!["250ms"]
    ),
    flag_float64_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { ratio := flag.Float64(\"ratio\", 0.5, \"\"); fmt.Println(*ratio) }",
        vec!["0.5"]
    ),
    flag_float64_after_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { ratio := flag.Float64(\"ratio\", 0.0, \"\"); _ = flag.Set(\"ratio\", \"1.25\"); fmt.Println(*ratio) }",
        vec!["1.25"]
    ),
    flag_string_set_overrides_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { mode := flag.String(\"mode\", \"dev\", \"\"); _ = flag.Set(\"mode\", \"prod\"); fmt.Println(*mode) }",
        vec!["prod"]
    ),
    flag_bool_set_to_true => (
        "package main; import \"fmt\"; import \"flag\"; func main() { verbose := flag.Bool(\"verbose\", false, \"\"); _ = flag.Set(\"verbose\", \"true\"); fmt.Println(*verbose) }",
        vec!["true"]
    ),
    flag_bool_set_to_false => (
        "package main; import \"fmt\"; import \"flag\"; func main() { debug := flag.Bool(\"debug\", true, \"\"); _ = flag.Set(\"debug\", \"false\"); fmt.Println(*debug) }",
        vec!["false"]
    ),
    flag_int_set_negative => (
        "package main; import \"fmt\"; import \"flag\"; func main() { offset := flag.Int(\"offset\", 0, \"\"); _ = flag.Set(\"offset\", \"-3\"); fmt.Println(*offset) }",
        vec!["-3"]
    ),
    flag_lookup_found_after_define => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.String(\"color\", \"red\", \"\"); f := flag.Lookup(\"color\"); fmt.Println(f != nil) }",
        vec!["true"]
    ),
    flag_lookup_missing_returns_nil => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fmt.Println(flag.Lookup(\"missing\") == nil) }",
        vec!["true"]
    ),
    flag_lookup_name_matches => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.Int(\"port\", 80, \"\"); f := flag.Lookup(\"port\"); fmt.Println(f.Name()) }",
        vec!["port"]
    ),
    flag_visit_all_counts_definitions => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.String(\"a\", \"\", \"\"); _ = flag.Int(\"b\", 0, \"\"); n := 0; flag.VisitAll(func(f *flag.Flag) { n++ }); fmt.Println(n >= 2) }",
        vec!["true"]
    ),
    flag_visit_all_collects_names => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.Bool(\"alpha\", false, \"\"); _ = flag.Bool(\"beta\", false, \"\"); found := 0; flag.VisitAll(func(f *flag.Flag) { if f.Name() == \"alpha\" || f.Name() == \"beta\" { found++ } }); fmt.Println(found) }",
        vec!["2"]
    ),
    flag_new_flagset_string_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); name := fs.String(\"name\", \"anon\", \"\"); fmt.Println(*name) }",
        vec!["anon"]
    ),
    flag_new_flagset_int_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); n := fs.Int(\"n\", 7, \"\"); fmt.Println(*n) }",
        vec!["7"]
    ),
    flag_new_flagset_set_isolated => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); mode := fs.String(\"mode\", \"dev\", \"\"); _ = fs.Set(\"mode\", \"prod\"); fmt.Println(*mode) }",
        vec!["prod"]
    ),
    flag_new_flagset_lookup_isolated => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.String(\"x\", \"1\", \"\"); fmt.Println(fs.Lookup(\"x\") != nil) }",
        vec!["true"]
    ),
    flag_commandline_set_does_not_affect_new_flagset => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.String(\"shared\", \"cmd\", \"\"); fs := flag.NewFlagSet(\"other\", flag.ContinueOnError); local := fs.String(\"shared\", \"local\", \"\"); fmt.Println(*local) }",
        vec!["local"]
    ),
    flag_int64_var_package_scope => (
        "package main; import \"fmt\"; import \"flag\"; var limit = flag.Int64(\"limit\", 50, \"\"); func main() { fmt.Println(*limit) }",
        vec!["50"]
    ),
    flag_uint_var_package_scope => (
        "package main; import \"fmt\"; import \"flag\"; var shards = flag.Uint(\"shards\", 3, \"\"); func main() { fmt.Println(*shards) }",
        vec!["3"]
    ),
    flag_duration_var_package_scope => (
        "package main; import \"fmt\"; import \"flag\"; import \"time\"; var delay = flag.Duration(\"delay\", time.Millisecond, \"\"); func main() { fmt.Println(*delay) }",
        vec!["1ms"]
    ),
    flag_multiple_types_independent_defaults => (
        "package main; import \"fmt\"; import \"flag\"; func main() { s := flag.String(\"s\", \"a\", \"\"); i := flag.Int(\"i\", 1, \"\"); b := flag.Bool(\"b\", true, \"\"); fmt.Println(*s, *i, *b) }",
        vec!["a 1 true"]
    ),
    flag_set_int64_string => (
        "package main; import \"fmt\"; import \"flag\"; func main() { v := flag.Int64(\"v\", 0, \"\"); _ = flag.Set(\"v\", \"9223372036854775807\"); fmt.Println(*v > 0) }",
        vec!["true"]
    ),
    flag_set_uint_max => (
        "package main; import \"fmt\"; import \"flag\"; func main() { v := flag.Uint(\"v\", 0, \"\"); _ = flag.Set(\"v\", \"4294967295\"); fmt.Println(*v) }",
        vec!["4294967295"]
    ),
    flag_new_flagset_visit_all_empty => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"empty\", flag.ContinueOnError); n := 0; fs.VisitAll(func(f *flag.Flag) { n++ }); fmt.Println(n) }",
        vec!["0"]
    ),
    flag_new_flagset_visit_all_after_define => (
        "package main; import \"fmt\"; import \"flag\"; func main() { fs := flag.NewFlagSet(\"fs\", flag.ContinueOnError); _ = fs.Bool(\"on\", false, \"\"); n := 0; fs.VisitAll(func(f *flag.Flag) { n++ }); fmt.Println(n) }",
        vec!["1"]
    ),
    flag_lookup_default_value_string => (
        "package main; import \"fmt\"; import \"flag\"; func main() { _ = flag.String(\"region\", \"eu\", \"\"); f := flag.Lookup(\"region\"); fmt.Println(f.DefValue) }",
        vec!["eu"]
    ),
    flag_int_rebind_pointer_after_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { n := flag.Int(\"n\", 1, \"\"); _ = flag.Set(\"n\", \"8\"); *n = 9; fmt.Println(*n) }",
        vec!["9"]
    ),
    flag_duration_composite_set => (
        "package main; import \"fmt\"; import \"flag\"; func main() { d := flag.Duration(\"d\", 0, \"\"); _ = flag.Set(\"d\", \"1h30m\"); fmt.Println(*d) }",
        vec!["1h30m0s"]
    ),
    flag_string_empty_default => (
        "package main; import \"fmt\"; import \"flag\"; func main() { s := flag.String(\"s\", \"\", \"\"); fmt.Println(len(*s)) }",
        vec!["0"]
    ),
    flag_bool_set_literal_1 => (
        "package main; import \"fmt\"; import \"flag\"; func main() { b := flag.Bool(\"b\", false, \"\"); _ = flag.Set(\"b\", \"1\"); fmt.Println(*b) }",
        vec!["true"]
    ),
    flag_bool_set_literal_0 => (
        "package main; import \"fmt\"; import \"flag\"; func main() { b := flag.Bool(\"b\", true, \"\"); _ = flag.Set(\"b\", \"0\"); fmt.Println(*b) }",
        vec!["false"]
    ),
}

go_compile_cases! {
    flag_parse_after_int64_uint_duration => "package main; import \"flag\"; func main() { _ = flag.Int64(\"size\", 0, \"\"); _ = flag.Uint(\"count\", 0, \"\"); _ = flag.Duration(\"timeout\", 0, \"\"); flag.Parse() }",
    flag_new_flagset_exit_on_error => "package main; import \"flag\"; func main() { _ = flag.NewFlagSet(\"app\", flag.ExitOnError) }",
    flag_new_flagset_panic_on_error => "package main; import \"flag\"; func main() { _ = flag.NewFlagSet(\"app\", flag.PanicOnError) }",
    flag_continue_on_error_constant => "package main; import \"flag\"; func main() { _ = flag.ContinueOnError }",
    flag_exit_on_error_constant => "package main; import \"flag\"; func main() { _ = flag.ExitOnError }",
    flag_panic_on_error_constant => "package main; import \"flag\"; func main() { _ = flag.PanicOnError }",
    flag_new_flagset_parse_nil_args => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Parse(nil) }",
    flag_new_flagset_parse_empty_slice => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Parse([]string{}) }",
    flag_new_flagset_lookup_missing => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Lookup(\"none\") }",
    flag_new_flagset_set_unknown => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Set(\"missing\", \"x\") }",
    flag_new_flagset_int64 => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Int64(\"n\", 0, \"\") }",
    flag_new_flagset_uint => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Uint(\"n\", 0, \"\") }",
    flag_new_flagset_uint64 => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Uint64(\"n\", 0, \"\") }",
    flag_new_flagset_duration => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Duration(\"d\", 0, \"\") }",
    flag_new_flagset_float64 => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Float64(\"f\", 0, \"\") }",
    flag_new_flagset_bool => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.Bool(\"b\", false, \"\") }",
    flag_new_flagset_string => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); _ = fs.String(\"s\", \"\", \"\") }",
    flag_new_flagset_visit_all => "package main; import \"flag\"; func main() { fs := flag.NewFlagSet(\"tool\", flag.ContinueOnError); fs.VisitAll(func(f *flag.Flag) {}) }",
    flag_commandline_vs_new_flagset_names => "package main; import \"flag\"; func main() { _ = flag.CommandLine; fs := flag.NewFlagSet(\"sub\", flag.ContinueOnError); _ = fs }",
    flag_unquote_usage_compile => "package main; import \"flag\"; func main() { _ = flag.UnquoteUsage }",
    flag_nflag_after_set => "package main; import \"flag\"; func main() { _ = flag.Bool(\"v\", false, \"\"); _ = flag.Set(\"v\", \"true\"); _ = flag.NFlag() }",
    flag_args_default_empty => "package main; import \"flag\"; func main() { _ = flag.Args() }",
    flag_narg_default_zero => "package main; import \"flag\"; func main() { _ = flag.NArg() }",
}
