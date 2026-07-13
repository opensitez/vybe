use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn pointer_incomplete_struct() { assert_eq!(run_c("struct Incomplete; int main() { struct Incomplete *p = 0; printf(\"%d\", p == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_array() { assert_eq!(run_c("int main() { int arr[]; int (*p)[] = &arr; printf(\"ok\"); return 0; } int arr[5];"), vec!["ok"]); }
#[test] fn pointer_incomplete_deref_fails() { assert_eq!(run_c("/* struct Incomplete; int main() { struct Incomplete *p = 0; *p; return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Cannot deref incomplete type
#[test] fn pointer_incomplete_arithmetic_fails() { assert_eq!(run_c("/* struct Incomplete; int main() { struct Incomplete *p = 0; p++; return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Cannot do arithmetic on incomplete type
#[test] fn pointer_incomplete_sizeof_fails() { assert_eq!(run_c("/* struct Incomplete; int main() { struct Incomplete *p; int s = sizeof(*p); return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // sizeof incomplete type is invalid
#[test] fn pointer_incomplete_cast_to_void() { assert_eq!(run_c("struct Incomplete; int main() { struct Incomplete *p = 0; void *v = p; printf(\"%d\", v == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_cast_from_void() { assert_eq!(run_c("struct Incomplete; int main() { void *v = 0; struct Incomplete *p = v; printf(\"%d\", p == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_comparison() { assert_eq!(run_c("struct Incomplete; int main() { struct Incomplete *p1 = 0; struct Incomplete *p2 = 0; printf(\"%d\", p1 == p2); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_in_struct() { assert_eq!(run_c("struct Incomplete; struct Container { struct Incomplete *p; }; int main() { struct Container c = {0}; printf(\"%d\", c.p == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_function_arg() { assert_eq!(run_c("struct Incomplete; void f(struct Incomplete *p) {} int main() { f(0); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn pointer_incomplete_function_return() { assert_eq!(run_c("struct Incomplete; struct Incomplete *f() { return 0; } int main() { printf(\"%d\", f() == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_typedef() { assert_eq!(run_c("typedef struct Incomplete Inc; int main() { Inc *p = 0; printf(\"%d\", p == 0); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_self_referential() { assert_eq!(run_c("struct Node { struct Node *next; }; int main() { struct Node n; n.next = &n; printf(\"%d\", n.next == &n); return 0; }"), vec!["1"]); }
#[test] fn pointer_incomplete_completion() { assert_eq!(run_c("struct Incomplete; struct Incomplete *p; struct Incomplete { int x; }; int main() { struct Incomplete i = {5}; p = &i; printf(\"%d\", p->x); return 0; }"), vec!["5"]); }
