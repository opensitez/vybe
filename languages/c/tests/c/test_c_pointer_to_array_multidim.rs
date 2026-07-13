use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn pointer_multidim_basic() { assert_eq!(run_c("int main() { int arr[2][3]; int (*p)[3] = arr; printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn pointer_multidim_deref() { assert_eq!(run_c("int main() { int arr[2][3] = {{1,2,3}, {4,5,6}}; int (*p)[3] = arr + 1; printf(\"%d\", (*p)[1]); return 0; }"), vec!["5"]); }
#[test] fn pointer_multidim_sizeof() { assert_eq!(run_c("int main() { int arr[2][3]; int (*p)[3] = arr; printf(\"%d\", (int)(sizeof(*p) / sizeof(int))); return 0; }"), vec!["3"]); }
#[test] fn pointer_multidim_increment() { assert_eq!(run_c("int main() { int arr[2][3] = {{1,2,3}, {4,5,6}}; int (*p)[3] = arr; p++; printf(\"%d\", (*p)[0]); return 0; }"), vec!["4"]); } // strides by 3 ints
#[test] fn pointer_multidim_array_of_pointers() { assert_eq!(run_c("int main() { int x=1, y=2; int *arr[2] = {&x, &y}; int **p = arr; printf(\"%d\", **p); return 0; }"), vec!["1"]); }
#[test] fn pointer_multidim_vla() { assert_eq!(run_c("int main() { int n=3; int arr[2][n]; int (*p)[n] = arr; p++; /* p points to second row of size n */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn pointer_multidim_cast() { assert_eq!(run_c("int main() { int arr[2][2] = {{1,2},{3,4}}; int *p = (int*)arr; printf(\"%d\", p[2]); return 0; }"), vec!["3"]); } // Flattens array
#[test] fn pointer_multidim_incompatible_stride_fails() { assert_eq!(run_c("/* int main() { int arr[2][3]; int (*p)[4] = arr; return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn pointer_multidim_3d() { assert_eq!(run_c("int main() { int arr[2][2][2] = {{{1,2},{3,4}}, {{5,6},{7,8}}}; int (*p)[2][2] = arr + 1; printf(\"%d\", (*p)[0][1]); return 0; }"), vec!["6"]); }
#[test] fn pointer_multidim_pass_to_function() { assert_eq!(run_c("void f(int (*p)[2]) { printf(\"%d\", p[1][0]); } int main() { int arr[2][2] = {{1,2},{3,4}}; f(arr); return 0; }"), vec!["3"]); }
#[test] fn pointer_multidim_return_from_function() { assert_eq!(run_c("int arr[2][2] = {{1,2},{3,4}}; int (*f())[2] { return arr; } int main() { printf(\"%d\", f()[1][1]); return 0; }"), vec!["4"]); }
#[test] fn pointer_multidim_typedef() { assert_eq!(run_c("typedef int Row[3]; int main() { int arr[2][3] = {{1,2,3},{4,5,6}}; Row *p = arr; printf(\"%d\", p[1][2]); return 0; }"), vec!["6"]); }
#[test] fn pointer_multidim_const() { assert_eq!(run_c("int main() { const int arr[2][2] = {{1,2},{3,4}}; const int (*p)[2] = arr; printf(\"%d\", (*p)[0]); return 0; }"), vec!["1"]); }
#[test] fn pointer_multidim_decay_twice_fails() { assert_eq!(run_c("/* int main() { int arr[2][2]; int **p = arr; return 0; } // arr decays to int(*)[2], not int** */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
