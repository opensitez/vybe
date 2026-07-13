use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn builtin_ffs() { assert_eq!(run_c("int main() { printf(\"%d %d\", __builtin_ffs(0), __builtin_ffs(8)); return 0; }"), vec!["0 4"]); }
#[test] fn builtin_clz() { assert_eq!(run_c("int main() { printf(\"%d\", __builtin_clz(1 << 30) >= 1); return 0; }"), vec!["1"]); } // Value is arch dependent, but > 0
#[test] fn builtin_ctz() { assert_eq!(run_c("int main() { printf(\"%d\", __builtin_ctz(8)); return 0; }"), vec!["3"]); }
#[test] fn builtin_popcount() { assert_eq!(run_c("int main() { printf(\"%d\", __builtin_popcount(0x11)); return 0; }"), vec!["2"]); }
#[test] fn builtin_parity() { assert_eq!(run_c("int main() { printf(\"%d %d\", __builtin_parity(0x3), __builtin_parity(0x7)); return 0; }"), vec!["0 1"]); }
#[test] fn builtin_expect() { assert_eq!(run_c("int main() { if (__builtin_expect(1 == 1, 1)) printf(\"expected\"); return 0; }"), vec!["expected"]); }
#[test] fn builtin_unreachable() { assert_eq!(run_c("int main() { if (0) __builtin_unreachable(); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn builtin_types_compatible_p() { assert_eq!(run_c("int main() { printf(\"%d %d\", __builtin_types_compatible_p(int, int), __builtin_types_compatible_p(int, float)); return 0; }"), vec!["1 0"]); }
#[test] fn builtin_constant_p() { assert_eq!(run_c("int main() { int x = 5; printf(\"%d %d\", __builtin_constant_p(42), __builtin_constant_p(x)); return 0; }"), vec!["1 0"]); }
#[test] fn builtin_add_overflow() { assert_eq!(run_c("int main() { int res; _Bool overflow = __builtin_add_overflow(2147483647, 1, &res); printf(\"%d\", overflow); return 0; }"), vec!["1"]); }
#[test] fn builtin_sub_overflow() { assert_eq!(run_c("int main() { int res; _Bool overflow = __builtin_sub_overflow(-2147483647, 2, &res); printf(\"%d\", overflow); return 0; }"), vec!["1"]); }
#[test] fn builtin_mul_overflow() { assert_eq!(run_c("int main() { int res; _Bool overflow = __builtin_mul_overflow(1000000, 1000000, &res); printf(\"%d\", overflow); return 0; }"), vec!["1"]); }
#[test] fn builtin_clzl() { assert_eq!(run_c("int main() { printf(\"%d\", __builtin_clzl(1L << 30) >= 1); return 0; }"), vec!["1"]); }
#[test] fn builtin_ctzll() { assert_eq!(run_c("int main() { printf(\"%d\", __builtin_ctzll(16LL)); return 0; }"), vec!["4"]); }
#[test] fn builtin_bswap16() { assert_eq!(run_c("int main() { printf(\"%x\", __builtin_bswap16(0x1234)); return 0; }"), vec!["3412"]); }
#[test] fn builtin_bswap32() { assert_eq!(run_c("int main() { printf(\"%x\", __builtin_bswap32(0x12345678)); return 0; }"), vec!["78563412"]); }
#[test] fn builtin_bswap64() { assert_eq!(run_c("int main() { printf(\"%llx\", __builtin_bswap64(0x1122334455667788ULL)); return 0; }"), vec!["8877665544332211"]); }
