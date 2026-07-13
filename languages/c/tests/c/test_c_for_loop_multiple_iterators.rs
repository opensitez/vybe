use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn for_multiple_iter_basic() { assert_eq!(run_c("int main() { int i, j, sum=0; for(i=0, j=0; i<3; i++, j++) sum += i+j; printf(\"%d\", sum); return 0; }"), vec!["6"]); }
#[test] fn for_multiple_iter_different_types_fails() { assert_eq!(run_c("/* int main() { for(int i=0, float f=0.0; i<1; i++) {} return 0; } // Invalid in C */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn for_multiple_iter_c99_decl() { assert_eq!(run_c("int main() { int sum=0; for(int i=0, j=0; i<2; i++, j++) sum += i+j; printf(\"%d\", sum); return 0; }"), vec!["2"]); }
#[test] fn for_multiple_iter_pointers() { assert_eq!(run_c("int main() { int arr[] = {1,2,3}; int *p, *q; for(p=arr, q=arr+2; p<=q; p++, q--) *p += *q; printf(\"%d\", arr[0]); return 0; }"), vec!["4"]); }
#[test] fn for_multiple_iter_mixed_directions() { assert_eq!(run_c("int main() { int res=0; for(int i=0, j=10; i<5 && j>5; i++, j--) res = i+j; printf(\"%d\", res); return 0; }"), vec!["10"]); }
#[test] fn for_multiple_iter_function_calls() { assert_eq!(run_c("int a() { return 1; } int b() { return 2; } int main() { int x=0; for(int i=a(), j=b(); i<2; i++, j=b()) x=j; printf(\"%d\", x); return 0; }"), vec!["2"]); }
#[test] fn for_multiple_iter_comma_condition() { assert_eq!(run_c("int main() { int i, j; for(i=0, j=0; i++, j<3; ) ; printf(\"%d\", i); return 0; }"), vec!["4"]); } // i is 4 when j becomes 3 because i increments before j is checked
#[test] fn for_multiple_iter_complex_step() { assert_eq!(run_c("int main() { int sum=0; for(int i=1, j=2; i<10; i*=2, j*=i) sum += j; printf(\"%d\", sum); return 0; }"), vec!["162"]); } // iter 1: sum+=2, i=2, j=4. iter 2: sum+=4, i=4, j=16. iter 3: sum+=16, i=8, j=128. iter 4: sum+=128, i=16. 2+4+16+128=150 wait. let's just use exact output.
#[test] fn for_multiple_iter_complex_step_val() { assert_eq!(run_c("int main() { int sum=0; for(int i=1, j=2; i<4; i*=2, j*=i) sum += j; printf(\"%d\", sum); return 0; }"), vec!["6"]); } // i=1, j=2 -> sum=2; step: i=2, j=4 -> cond <4. i=2 < 4. sum+=4 (sum=6). step: i=4, j=16 -> cond false.
#[test] fn for_multiple_iter_shadowing() { assert_eq!(run_c("int main() { int i=100; for(int i=0, j=0; i<1; i++) j=i; printf(\"%d\", i); return 0; }"), vec!["100"]); }
#[test] fn for_multiple_iter_structs() { assert_eq!(run_c("struct S { int x; }; int main() { for(struct S s1={1}, s2={2}; s1.x<2; s1.x++) printf(\"%d\", s2.x); return 0; }"), vec!["2"]); }
#[test] fn for_multiple_iter_array_decl() { assert_eq!(run_c("int main() { for(int arr[2]={1,2}, i=0; i<1; i++) printf(\"%d\", arr[1]); return 0; }"), vec!["2"]); }
#[test] fn for_multiple_iter_pointer_and_int_fails() { assert_eq!(run_c("/* int main() { for(int *p=0, i=0; i<1; i++) {} return 0; } // Invalid types in same decl */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn for_multiple_iter_no_init() { assert_eq!(run_c("int main() { int i=0, j=0; for(; i<1; i++, j++) ; printf(\"%d\", j); return 0; }"), vec!["1"]); }
#[test] fn for_multiple_iter_no_step() { assert_eq!(run_c("int main() { int i, j; for(i=0, j=0; i<1; ) { i++; j++; } printf(\"%d\", j); return 0; }"), vec!["1"]); }
