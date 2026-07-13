use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn printf_prec_string_truncate() { assert_eq!(run_c("int main() { printf(\"%.3s\", \"hello\"); return 0; }"), vec!["hel"]); }
#[test] fn printf_prec_string_no_truncate() { assert_eq!(run_c("int main() { printf(\"%.10s\", \"hello\"); return 0; }"), vec!["hello"]); }
#[test] fn printf_prec_int_zero_pad() { assert_eq!(run_c("int main() { printf(\"%.5d\", 42); return 0; }"), vec!["00042"]); }
#[test] fn printf_prec_int_no_pad() { assert_eq!(run_c("int main() { printf(\"%.1d\", 42); return 0; }"), vec!["42"]); }
#[test] fn printf_prec_int_zero_value_zero_prec() { assert_eq!(run_c("int main() { printf(\"|%.0d|\", 0); return 0; }"), vec!["||"]); } // empty
#[test] fn printf_prec_float_rounding() { assert_eq!(run_c("int main() { printf(\"%.2f\", 3.14159); return 0; }"), vec!["3.14"]); }
#[test] fn printf_prec_float_zero_prec() { assert_eq!(run_c("int main() { printf(\"%.0f\", 3.14159); return 0; }"), vec!["3"]); }
#[test] fn printf_prec_dynamic() { assert_eq!(run_c("int main() { printf(\"%.*s\", 3, \"hello\"); return 0; }"), vec!["hel"]); }
#[test] fn printf_width_and_prec() { assert_eq!(run_c("int main() { printf(\"|%5.3s|\", \"hello\"); return 0; }"), vec!["|  hel|"]); } // truncate to 3, pad to 5
#[test] fn printf_width_and_prec_dynamic() { assert_eq!(run_c("int main() { printf(\"|%*.*s|\", 5, 3, \"hello\"); return 0; }"), vec!["|  hel|"]); }
#[test] fn printf_prec_dynamic_negative_ignored() { assert_eq!(run_c("int main() { printf(\"%.*f\", -2, 3.14159); return 0; }"), vec!["3.141590"]); } // negative prec treated as if missing
