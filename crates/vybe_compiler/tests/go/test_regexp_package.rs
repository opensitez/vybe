//! regexp package: Match, Find, Replace, Compile patterns.

use crate::helpers::*;

go_run_cases! {
    regexp_match_string_literal => ("package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.MatchString(\"^go\", \"gopher\")) }", vec!["true"]),
    regexp_match_string_miss => ("package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.MatchString(\"^rust\", \"gopher\")) }", vec!["false"]),
    regexp_find_first_submatch => ("package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)`); m := re.FindStringSubmatch(\"id:42\"); fmt.Println(m[1]) }", vec!["42"]),
    regexp_replace_all => ("package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`a+`); fmt.Println(re.ReplaceAllString(\"baaac\", \"X\")) }", vec!["bXc"]),
    regexp_split => ("package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`[,\\s]+`); parts := re.Split(\"a, b  c\", -1); fmt.Println(len(parts)); fmt.Println(parts[2]) }", vec!["3", "c"]),
}

go_compile_cases! {
    regexp_compile_anchor => "package main; import \"regexp\"; func main() { _, _ = regexp.Compile(`^start`) }",
    regexp_quote_meta_chars => "package main; import \"regexp\"; func main() { _ = regexp.QuoteMeta(\"a.b\") }",
    regexp_find_all_index => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`x`); _ = re.FindAllStringIndex(\"xxy\", -1) }",
    regexp_num_subexp => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(a)(b)`); _ = re.NumSubexp() }",
    regexp_literal_metachar => "package main; import \"regexp\"; func main() { _ = regexp.MustCompile(`\\d+`) }",
}
