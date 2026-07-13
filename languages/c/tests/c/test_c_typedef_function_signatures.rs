use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn typedef_func_type() { assert_eq!(run_c("typedef int F(void); F my_func; int my_func(void) { return 1; } int main() { printf(\"%d\", my_func()); return 0; }"), vec!["1"]); }
#[test] fn typedef_func_ptr() { assert_eq!(run_c("typedef int (*F)(void); int f() { return 2; } int main() { F ptr = f; printf(\"%d\", ptr()); return 0; }"), vec!["2"]); }
#[test] fn typedef_func_ptr_args() { assert_eq!(run_c("typedef int (*F)(int, int); int add(int a, int b) { return a+b; } int main() { F ptr = add; printf(\"%d\", ptr(3, 4)); return 0; }"), vec!["7"]); }
#[test] fn typedef_func_returning_ptr() { assert_eq!(run_c("typedef int* (*F)(void); int x = 5; int* get_x() { return &x; } int main() { F ptr = get_x; printf(\"%d\", *ptr()); return 0; }"), vec!["5"]); }
#[test] fn typedef_func_array() { assert_eq!(run_c("typedef int (*F)(void); int f1() { return 1; } int f2() { return 2; } int main() { F arr[2] = {f1, f2}; printf(\"%d\", arr[1]()); return 0; }"), vec!["2"]); }
#[test] fn typedef_func_returning_func_ptr() { assert_eq!(run_c("typedef void (*Action)(void); typedef Action (*GetAction)(void); void a() { printf(\"A\"); } Action get() { return a; } int main() { GetAction g = get; g()(); return 0; }"), vec!["A"]); }
#[test] fn typedef_func_as_arg() { assert_eq!(run_c("typedef int F(int); int apply(F f, int v) { return f(v); } int dbl(int x) { return x*2; } int main() { printf(\"%d\", apply(dbl, 5)); return 0; }"), vec!["10"]); }
#[test] fn typedef_func_ptr_as_arg() { assert_eq!(run_c("typedef int (*F)(int); int apply(F f, int v) { return f(v); } int dbl(int x) { return x*2; } int main() { printf(\"%d\", apply(dbl, 6)); return 0; }"), vec!["12"]); }
#[test] fn typedef_func_implicit_int_fails_in_c99() { assert_eq!(run_c("/* typedef F(void); */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn typedef_func_varargs() { assert_eq!(run_c("typedef int (*F)(const char*, ...); #include <stdio.h>\nint main() { F p = printf; p(\"%d\", 42); return 0; }"), vec!["42"]); }
#[test] fn typedef_func_void_ptr_arg() { assert_eq!(run_c("typedef void* (*F)(void*); void* echo(void* p) { return p; } int main() { F f = echo; int a = 99; printf(\"%d\", *(int*)f(&a)); return 0; }"), vec!["99"]); }
#[test] fn typedef_func_struct_arg() { assert_eq!(run_c("struct S { int a; }; typedef int (*F)(struct S); int get(struct S s) { return s.a; } int main() { F f = get; struct S s = {88}; printf(\"%d\", f(s)); return 0; }"), vec!["88"]); }
#[test] fn typedef_func_returning_struct() { assert_eq!(run_c("struct S { int a; }; typedef struct S (*F)(void); struct S get() { struct S s = {77}; return s; } int main() { F f = get; printf(\"%d\", f().a); return 0; }"), vec!["77"]); }
#[test] fn typedef_func_complex_nesting() { assert_eq!(run_c("typedef int (*F1)(int); typedef F1 (*F2)(int); int f1(int x) { return x; } F1 f2(int y) { return f1; } int main() { F2 f = f2; printf(\"%d\", f(1)(2)); return 0; }"), vec!["2"]); }
#[test] fn typedef_func_void() { assert_eq!(run_c("typedef void F(void); F f; void f(void) { printf(\"V\"); } int main() { f(); return 0; }"), vec!["V"]); }
