//! log/slog: Level handling, LogAttrs, With/Group, TextHandler options,
//! value types, Default logger — extended coverage distinct from
//! `test_cover_text_html_log.rs` compile smokes.

use crate::helpers::*;

go_run_cases! {
    slog_text_handler_info_message_present => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; h := slog.NewTextHandler(&buf, nil); l := slog.New(h); l.Info(\"ready\"); fmt.Println(strings.Contains(buf.String(), \"ready\")) }",
        vec!["true"]
    ),
    slog_text_handler_debug_message_present => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; h := slog.NewTextHandler(&buf, nil); l := slog.New(h); l.Debug(\"trace\"); fmt.Println(strings.Contains(buf.String(), \"trace\")) }",
        vec!["true"]
    ),
    slog_int_attr_in_text_output => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"n\", slog.Int(\"count\", 7)); s := buf.String(); fmt.Println(strings.Contains(s, \"count\")); fmt.Println(strings.Contains(s, \"7\")) }",
        vec!["true", "true"]
    ),
    slog_string_attr_in_text_output => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"msg\", slog.String(\"role\", \"admin\")); fmt.Println(strings.Contains(buf.String(), \"admin\")) }",
        vec!["true"]
    ),
    slog_bool_attr_true_in_text_output => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"flag\", slog.Bool(\"ok\", true)); fmt.Println(strings.Contains(buf.String(), \"true\")) }",
        vec!["true"]
    ),
    slog_bool_attr_false_in_text_output => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"flag\", slog.Bool(\"ok\", false)); fmt.Println(strings.Contains(buf.String(), \"false\")) }",
        vec!["true"]
    ),
    slog_any_attr_int_in_text_output => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"v\", slog.Any(\"x\", 42)); fmt.Println(strings.Contains(buf.String(), \"42\")) }",
        vec!["true"]
    ),
    slog_with_prepends_constant_attr => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; base := slog.New(slog.NewTextHandler(&buf, nil)); child := base.With(\"svc\", \"api\"); child.Info(\"hit\"); fmt.Println(strings.Contains(buf.String(), \"api\")) }",
        vec!["true"]
    ),
    slog_with_group_prefixes_key => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; base := slog.New(slog.NewTextHandler(&buf, nil)); child := base.WithGroup(\"req\"); child.Info(\"in\", slog.Int(\"id\", 1)); fmt.Println(strings.Contains(buf.String(), \"req\")) }",
        vec!["true"]
    ),
    slog_group_value_expands_nested_attrs => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"evt\", slog.Group(\"meta\", slog.String(\"env\", \"dev\"))); fmt.Println(strings.Contains(buf.String(), \"dev\")) }",
        vec!["true"]
    ),
    slog_level_info_string_label => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelInfo.String()) }",
        vec!["INFO"]
    ),
    slog_level_debug_string_label => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelDebug.String()) }",
        vec!["DEBUG"]
    ),
    slog_level_warn_string_label => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelWarn.String()) }",
        vec!["WARN"]
    ),
    slog_level_error_string_label => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelError.String()) }",
        vec!["ERROR"]
    ),
    slog_default_info_writes_message => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { slog.Default().Info(\"default\"); fmt.Println(\"ok\") }",
        vec!["ok"]
    ),
    slog_log_attrs_int_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"context\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.LogAttrs(context.Background(), slog.LevelInfo, \"attrs\", slog.Int(\"n\", 3)); fmt.Println(strings.Contains(buf.String(), \"3\")) }",
        vec!["true"]
    ),
    slog_log_attrs_string_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"context\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.LogAttrs(context.Background(), slog.LevelInfo, \"attrs\", slog.String(\"k\", \"v\")); fmt.Println(strings.Contains(buf.String(), \"v\")) }",
        vec!["true"]
    ),
    slog_log_attrs_bool_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"context\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.LogAttrs(context.Background(), slog.LevelInfo, \"attrs\", slog.Bool(\"on\", true)); fmt.Println(strings.Contains(buf.String(), \"true\")) }",
        vec!["true"]
    ),
    slog_log_attrs_group_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"context\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.LogAttrs(context.Background(), slog.LevelInfo, \"evt\", slog.Group(\"g\", slog.Int(\"n\", 9))); fmt.Println(strings.Contains(buf.String(), \"9\")) }",
        vec!["true"]
    ),
    slog_text_handler_add_source_option => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; opts := &slog.HandlerOptions{AddSource: true}; h := slog.NewTextHandler(&buf, opts); l := slog.New(h); l.Info(\"src\"); fmt.Println(len(buf.String()) > 0) }",
        vec!["true"]
    ),
    slog_text_handler_level_filter_debug => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; opts := &slog.HandlerOptions{Level: slog.LevelDebug}; h := slog.NewTextHandler(&buf, opts); l := slog.New(h); l.Debug(\"low\"); fmt.Println(buf.Len() > 0) }",
        vec!["true"]
    ),
    slog_text_handler_level_filter_error => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; opts := &slog.HandlerOptions{Level: slog.LevelError}; h := slog.NewTextHandler(&buf, opts); l := slog.New(h); l.Info(\"hidden\"); fmt.Println(buf.Len() == 0) }",
        vec!["true"]
    ),
    slog_with_two_attrs_both_present => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)).With(\"a\", 1, \"b\", 2); l.Info(\"pair\"); s := buf.String(); fmt.Println(strings.Contains(s, \"1\")); fmt.Println(strings.Contains(s, \"2\")) }",
        vec!["true", "true"]
    ),
    slog_int64_attr_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"big\", slog.Int64(\"n\", 100)); fmt.Println(strings.Contains(buf.String(), \"100\")) }",
        vec!["true"]
    ),
    slog_uint64_attr_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"u\", slog.Uint64(\"n\", 5)); fmt.Println(strings.Contains(buf.String(), \"5\")) }",
        vec!["true"]
    ),
    slog_float64_attr_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"f\", slog.Float64(\"x\", 1.5)); fmt.Println(strings.Contains(buf.String(), \"1.5\")) }",
        vec!["true"]
    ),
    slog_duration_attr_value => (
        "package main; import \"fmt\"; import \"log/slog\"; import \"bytes\"; import \"strings\"; import \"time\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewTextHandler(&buf, nil)); l.Info(\"d\", slog.Duration(\"wait\", time.Second)); fmt.Println(strings.Contains(buf.String(), \"1s\")) }",
        vec!["true"]
    ),
    slog_warn_level_string => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelWarn.String()) }",
        vec!["WARN"]
    ),
    slog_error_level_string => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { fmt.Println(slog.LevelError.String()) }",
        vec!["ERROR"]
    ),
    slog_default_debug_compile_smoke => (
        "package main; import \"fmt\"; import \"log/slog\"; func main() { slog.Default().Debug(\"d\"); fmt.Println(\"done\") }",
        vec!["done"]
    ),
}

go_compile_cases! {
    slog_json_handler_info => "package main; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewJSONHandler(&buf, nil)); l.Info(\"json\") }",
    slog_json_handler_with_attrs => "package main; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; l := slog.New(slog.NewJSONHandler(&buf, nil)); l.Info(\"json\", slog.Int(\"n\", 1)) }",
    slog_json_handler_options_level => "package main; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; opts := &slog.HandlerOptions{Level: slog.LevelWarn}; l := slog.New(slog.NewJSONHandler(&buf, opts)); l.Warn(\"warn\") }",
    slog_set_default_custom_logger => "package main; import \"log/slog\"; import \"bytes\"; func main() { var buf bytes.Buffer; slog.SetDefault(slog.New(slog.NewTextHandler(&buf, nil))); slog.Info(\"after\") }",
    slog_with_package_level => "package main; import \"log/slog\"; func main() { _ = slog.With(\"k\", \"v\").With(\"k2\", 2) }",
    slog_with_group_package_level => "package main; import \"log/slog\"; func main() { _ = slog.WithGroup(\"outer\").WithGroup(\"inner\") }",
    slog_log_context_background => "package main; import \"log/slog\"; import \"context\"; func main() { slog.Log(context.Background(), slog.LevelInfo, \"ctx\") }",
    slog_log_attrs_any => "package main; import \"log/slog\"; import \"context\"; func main() { slog.LogAttrs(context.Background(), slog.LevelDebug, \"any\", slog.Any(\"v\", \"s\")) }",
    slog_new_record_time => "package main; import \"log/slog\"; import \"time\"; func main() { _ = slog.NewRecord(time.Now(), slog.LevelInfo, \"rec\", 0) }",
    slog_handler_with_attrs_chain => "package main; import \"log/slog\"; import \"bytes\"; func main() { h := slog.NewTextHandler(bytes.NewBuffer(nil), nil); _ = h.WithAttrs([]slog.Attr{slog.Int(\"n\", 1)}) }",
    slog_handler_with_group_chain => "package main; import \"log/slog\"; import \"bytes\"; func main() { h := slog.NewTextHandler(bytes.NewBuffer(nil), nil); _ = h.WithGroup(\"g\") }",
    slog_level_enabled_default => "package main; import \"log/slog\"; import \"context\"; func main() { _ = slog.Default().Enabled(context.Background(), slog.LevelInfo) }",
    slog_set_log_logger_level_info => "package main; import \"log/slog\"; func main() { _ = slog.SetLogLoggerLevel(slog.LevelInfo) }",
    slog_text_handler_replace_attr => "package main; import \"log/slog\"; import \"bytes\"; func main() { opts := &slog.HandlerOptions{ReplaceAttr: func(groups []string, a slog.Attr) slog.Attr { return a }}; _ = slog.NewTextHandler(bytes.NewBuffer(nil), opts) }",
    slog_int_value_constructor => "package main; import \"log/slog\"; func main() { _ = slog.Int(\"k\", 0) }",
    slog_string_value_constructor => "package main; import \"log/slog\"; func main() { _ = slog.String(\"k\", \"\") }",
    slog_bool_value_constructor => "package main; import \"log/slog\"; func main() { _ = slog.Bool(\"k\", false) }",
    slog_any_value_constructor => "package main; import \"log/slog\"; func main() { _ = slog.Any(\"k\", nil) }",
    slog_group_value_constructor => "package main; import \"log/slog\"; func main() { _ = slog.Group(\"g\", slog.Bool(\"b\", true)) }",
    slog_default_warn => "package main; import \"log/slog\"; func main() { slog.Default().Warn(\"w\") }",
    slog_default_error => "package main; import \"log/slog\"; func main() { slog.Default().Error(\"e\") }",
    slog_log_level_error => "package main; import \"log/slog\"; import \"context\"; func main() { slog.Log(context.Background(), slog.LevelError, \"err\") }",
    slog_log_level_debug => "package main; import \"log/slog\"; import \"context\"; func main() { slog.Log(context.Background(), slog.LevelDebug, \"dbg\") }",
    slog_text_handler_nil_options => "package main; import \"log/slog\"; import \"bytes\"; func main() { _ = slog.NewTextHandler(&bytes.Buffer{}, nil) }",
    slog_json_handler_add_source => "package main; import \"log/slog\"; import \"bytes\"; func main() { opts := &slog.HandlerOptions{AddSource: true}; _ = slog.NewJSONHandler(&bytes.Buffer{}, opts) }",
}
