//! fmt.Sprintf / fmt.Print formatting verbs — distinct edit-descriptor semantics.

go_run_cases! {
    sprintf_decimal_positive => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%d\", 42)) }", vec!["42"]),
    sprintf_decimal_negative => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%d\", -7)) }", vec!["-7"]),
    sprintf_hex_lowercase => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%x\", 255)) }", vec!["ff"]),
    sprintf_hex_uppercase => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%X\", 255)) }", vec!["FF"]),
    sprintf_octal => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%o\", 8)) }", vec!["10"]),
    sprintf_binary => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%b\", 5)) }", vec!["101"]),
    sprintf_string => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%s\", \"go\")) }", vec!["go"]),
    sprintf_quoted_string => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%q\", \"go\")) }", vec!["\"go\""]),
    sprintf_bool_true => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%t\", true)) }", vec!["true"]),
    sprintf_bool_false => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%t\", false)) }", vec!["false"]),
    sprintf_default_verb_int => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", 9)) }", vec!["9"]),
    sprintf_default_verb_string => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", \"vy\")) }", vec!["vy"]),
    sprintf_float_fixed => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%.2f\", 1.234)) }", vec!["1.23"]),
    sprintf_float_scientific => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%e\", 1000.0)) }", vec!["1.000000e+03"]),
    sprintf_width_padded => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%5d\", 7)) }", vec!["    7"]),
    sprintf_plus_sign => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%+d\", 7)) }", vec!["+7"]),
    sprintf_multiple_values => ("package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%s-%d\", \"id\", 3)) }", vec!["id-3"]),
    sprintf_pointer_nil => ("package main; import \"fmt\"; func main() { var p *int; fmt.Println(fmt.Sprintf(\"%p\", p)) }", vec!["0x0"]),
}

go_compile_cases! {
    sprintf_go_string_verb => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%#v\", struct{ X int }{X: 1})) }",
    sprintf_go_string_int => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%#x\", 16)) }",
    sprintf_unicode_char => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%c\", 65)) }",
    sprintf_error_verb => "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", errors.New(\"boom\"))) }",
    sprintf_slice_brackets => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", []int{1,2})) }",
    sprintf_map_brackets => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", map[string]int{\"a\":1})) }",
    sprintf_width_asterisk => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%*d\", 4, 9)) }",
    sprintf_precision_asterisk => "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%.*f\", 1, 2.25)) }",
}
