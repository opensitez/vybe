use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn pointer_past_end_creation() { assert_eq!(run_c("int main() { int arr[3]; int *p = arr + 3; printf(\"%d\", p > arr); return 0; }"), vec!["1"]); } // Legal to point one past end
#[test] fn pointer_past_end_comparison() { assert_eq!(run_c("int main() { int arr[3] = {1,2,3}; int *p; int count=0; for(p=arr; p < arr+3; p++) count++; printf(\"%d\", count); return 0; }"), vec!["3"]); }
#[test] fn pointer_past_end_deref_fails() { assert_eq!(run_c("/* int main() { int arr[3]; int *p = arr + 3; *p = 5; return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // UB, don't execute
#[test] fn pointer_past_end_subtraction() { assert_eq!(run_c("int main() { int arr[3]; int *p1 = arr; int *p2 = arr + 3; printf(\"%d\", (int)(p2 - p1)); return 0; }"), vec!["3"]); }
#[test] fn pointer_before_beginning_fails() { assert_eq!(run_c("/* int main() { int arr[3]; int *p = arr - 1; return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // UB to even compute pointer before array
#[test] fn pointer_past_end_struct_array() { assert_eq!(run_c("struct S { int x; }; int main() { struct S arr[2]; struct S *p = arr + 2; printf(\"%d\", p > arr); return 0; }"), vec!["1"]); }
#[test] fn pointer_past_end_multidim() { assert_eq!(run_c("int main() { int arr[2][2]; int (*p)[2] = arr + 2; printf(\"%d\", p > arr); return 0; }"), vec!["1"]); }
#[test] fn pointer_past_end_scalar() { assert_eq!(run_c("int main() { int x; int *p = &x + 1; printf(\"%d\", p > &x); return 0; }"), vec!["1"]); } // Scalar acts like array of size 1
#[test] fn pointer_past_end_equality() { assert_eq!(run_c("int main() { int arr[2]; int *p = arr + 2; printf(\"%d\", p == &arr[2]); return 0; }"), vec!["1"]); } // &arr[2] is legal syntax for past end
#[test] fn pointer_past_end_loop_down() { assert_eq!(run_c("int main() { int arr[3] = {1,2,3}; int *p = arr + 3; int count=0; while(p > arr) { p--; count += *p; } printf(\"%d\", count); return 0; }"), vec!["6"]); }
#[test] fn pointer_past_end_vla() { assert_eq!(run_c("int main() { int n = 5; int arr[n]; int *p = arr + n; printf(\"%d\", (int)(p - arr)); return 0; }"), vec!["5"]); }
#[test] fn pointer_past_end_compound_literal() { assert_eq!(run_c("int main() { int *p = (int[]){1,2,3} + 3; printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn pointer_past_end_char_array() { assert_eq!(run_c("int main() { char str[] = \"abc\"; char *p = str + 4; printf(\"%d\", (int)(p - str)); return 0; }"), vec!["4"]); } // includes null terminator
#[test] fn pointer_past_end_zero_size_fails() { assert_eq!(run_c("/* int arr[0]; int *p = arr + 0; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // zero size arrays are non-standard
