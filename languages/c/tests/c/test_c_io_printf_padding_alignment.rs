use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn printf_right_pad() { assert_eq!(run_c("int main() { printf(\"|%5d|\", 42); return 0; }"), vec!["|   42|"]); }
#[test] fn printf_left_pad() { assert_eq!(run_c("int main() { printf(\"|%-5d|\", 42); return 0; }"), vec!["|42   |"]); }
#[test] fn printf_zero_pad() { assert_eq!(run_c("int main() { printf(\"|%05d|\", 42); return 0; }"), vec!["|00042|"]); }
#[test] fn printf_space_flag_positive() { assert_eq!(run_c("int main() { printf(\"|% d|\", 42); return 0; }"), vec!["| 42|"]); }
#[test] fn printf_space_flag_negative() { assert_eq!(run_c("int main() { printf(\"|% d|\", -42); return 0; }"), vec!["|-42|"]); }
#[test] fn printf_plus_flag_positive() { assert_eq!(run_c("int main() { printf(\"|%+d|\", 42); return 0; }"), vec!["|+42|"]); }
#[test] fn printf_plus_flag_negative() { assert_eq!(run_c("int main() { printf(\"|%+d|\", -42); return 0; }"), vec!["|-42|"]); }
#[test] fn printf_hash_flag_octal() { assert_eq!(run_c("int main() { printf(\"%#o\", 42); return 0; }"), vec!["052"]); }
#[test] fn printf_hash_flag_hex() { assert_eq!(run_c("int main() { printf(\"%#x\", 42); return 0; }"), vec!["0x2a"]); }
#[test] fn printf_hash_flag_hex_upper() { assert_eq!(run_c("int main() { printf(\"%#X\", 42); return 0; }"), vec!["0X2A"]); }
#[test] fn printf_hash_flag_float() { assert_eq!(run_c("int main() { printf(\"%#.0f\", 42.0); return 0; }"), vec!["42."]); } // forces decimal point
#[test] fn printf_dynamic_width() { assert_eq!(run_c("int main() { printf(\"|%*d|\", 5, 42); return 0; }"), vec!["|   42|"]); }
#[test] fn printf_dynamic_width_negative_is_left_pad() { assert_eq!(run_c("int main() { printf(\"|%*d|\", -5, 42); return 0; }"), vec!["|42   |"]); }
