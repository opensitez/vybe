use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn scanf_scanset_basic() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"abc123def\", \"%[a-z]\", buf); printf(\"%s\", buf); return 0; }"), vec!["abc"]); }
#[test] fn scanf_scanset_negated() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"abc123def\", \"%[^1-9]\", buf); printf(\"%s\", buf); return 0; }"), vec!["abc"]); }
#[test] fn scanf_scanset_include_bracket() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"]abc\", \"%[]a-z]\", buf); printf(\"%s\", buf); return 0; }"), vec!["]abc"]); } // ] must be first to be included
#[test] fn scanf_scanset_include_caret() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"^abc\", \"%[^]a-z]\", buf); printf(\"%s\", buf); return 0; }"), vec!["^"]); } // wait, this is tricky, let's just do a simple one
#[test] fn scanf_scanset_include_caret_simple() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"^abc\", \"%[a-z^]\", buf); printf(\"%s\", buf); return 0; }"), vec!["^abc"]); }
#[test] fn scanf_scanset_include_hyphen() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"a-b\", \"%[a-z-]\", buf); printf(\"%s\", buf); return 0; }"), vec!["a-b"]); } // hyphen must be last
#[test] fn scanf_scanset_width_limit() { assert_eq!(run_c("int main() { char buf[10]; sscanf(\"abcde\", \"%3[a-z]\", buf); printf(\"%s\", buf); return 0; }"), vec!["abc"]); }
#[test] fn scanf_scanset_empty_match() { assert_eq!(run_c("int main() { char buf[10] = \"xxx\"; int n = sscanf(\"123abc\", \"%[a-z]\", buf); printf(\"%d %s\", n, buf); return 0; }"), vec!["0 xxx"]); } // fails immediately
#[test] fn scanf_scanset_multiple() { assert_eq!(run_c("int main() { char b1[10], b2[10]; sscanf(\"abc123def\", \"%[a-z]%[0-9]\", b1, b2); printf(\"%s %s\", b1, b2); return 0; }"), vec!["abc 123"]); }
