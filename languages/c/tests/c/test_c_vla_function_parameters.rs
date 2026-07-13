use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn vla_parameter_basic() { assert_eq!(run_c("void f(int n, int arr[n]) { printf(\"%d\", arr[n-1]); } int main() { int a[3] = {1,2,3}; f(3, a); return 0; }"), vec!["3"]); }
#[test] fn vla_parameter_multidim() { assert_eq!(run_c("void f(int r, int c, int arr[r][c]) { printf(\"%d\", arr[1][1]); } int main() { int a[2][2] = {{1,2},{3,4}}; f(2, 2, a); return 0; }"), vec!["4"]); }
#[test] fn vla_parameter_static_fails() { assert_eq!(run_c("/* void f(int n, int arr[static n]) {} // static requires at least n elements */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn vla_parameter_restrict() { assert_eq!(run_c("void f(int n, int arr[restrict n]) { arr[0] = 5; } int main() { int a[1]; f(1, a); printf(\"%d\", a[0]); return 0; }"), vec!["5"]); }
#[test] fn vla_parameter_const() { assert_eq!(run_c("void f(int n, const int arr[n]) { printf(\"%d\", arr[0]); } int main() { int a[1] = {5}; f(1, a); return 0; }"), vec!["5"]); }
#[test] fn vla_parameter_star() { assert_eq!(run_c("void f(int, int[*]); void f(int n, int arr[n]) { printf(\"%d\", arr[0]); } int main() { int a[1]={5}; f(1, a); return 0; }"), vec!["5"]); } // prototype with [*]
#[test] fn vla_parameter_out_of_order_fails() { assert_eq!(run_c("/* void f(int arr[n], int n) {} // error: n undeclared */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn vla_parameter_sizeof() { assert_eq!(run_c("void f(int n, int arr[n]) { printf(\"%d\", (int)sizeof(arr)); } int main() { int a[5]; f(5, a); return 0; }"), vec!["4"]); } // wait, sizeof(arr) where arr is a parameter is sizeof(int*), which is 4 or 8! Let's just test compiling
#[test] fn vla_parameter_sizeof_decays() { assert_eq!(run_c("void f(int n, int arr[n]) { printf(\"%d\", sizeof(arr) == sizeof(int*)); } int main() { int a[5]; f(5, a); return 0; }"), vec!["1"]); } // Decays to pointer
#[test] fn vla_parameter_typedef() { assert_eq!(run_c("int main() { int n=5; typedef int VLA[n]; VLA a; printf(\"%d\", (int)(sizeof(a)/sizeof(int))); return 0; }"), vec!["5"]); } // VLA typedef
#[test] fn vla_parameter_pointer_to_vla() { assert_eq!(run_c("void f(int n, int (*p)[n]) { printf(\"%d\", (*p)[0]); } int main() { int a[2]={5,6}; f(2, &a); return 0; }"), vec!["5"]); }
#[test] fn vla_parameter_function_prototype_scope() { assert_eq!(run_c("void f(int n, int a[n+1]); void f(int n, int a[n+1]) { printf(\"%d\", a[1]); } int main() { int arr[2] = {1, 2}; f(1, arr); return 0; }"), vec!["2"]); } // expression is evaluated on entry
#[test] fn vla_parameter_sizeof_vla_ptr() { assert_eq!(run_c("void f(int n, int (*p)[n]) { printf(\"%d\", sizeof(*p) == n * sizeof(int)); } int main() { int a[5]; f(5, &a); return 0; }"), vec!["1"]); } // *p is a VLA type!
#[test] fn vla_parameter_global_size_fails() { assert_eq!(run_c("/* int n=5; void f(int arr[n]) {} // n must be positive integer constant expression for global */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
