use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn short_circuit_and_basic() { assert_eq!(run_c("int main() { int x=0; 0 && (x=1); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_or_basic() { assert_eq!(run_c("int main() { int x=0; 1 || (x=1); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_and_chain() { assert_eq!(run_c("int main() { int x=0, y=0; 1 && 0 && (x=1); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_or_chain() { assert_eq!(run_c("int main() { int x=0; 0 || 1 || (x=1); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_mixed_precedence_1() { assert_eq!(run_c("int main() { int x=0; 0 && 1 || (x=1); printf(\"%d\", x); return 0; }"), vec!["1"]); } // (0 && 1) is 0, so || evaluates RHS
#[test] fn short_circuit_mixed_precedence_2() { assert_eq!(run_c("int main() { int x=0; 1 || 0 && (x=1); printf(\"%d\", x); return 0; }"), vec!["0"]); } // 1 || (0 && (x=1)), 1 is true, RHS not evaluated
#[test] fn short_circuit_function_calls() { assert_eq!(run_c("int f(int *x) { *x += 1; return 1; } int main() { int x=0; 0 && f(&x); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_function_calls_evaluated() { assert_eq!(run_c("int f(int *x) { *x += 1; return 1; } int main() { int x=0; 1 && f(&x); printf(\"%d\", x); return 0; }"), vec!["1"]); }
#[test] fn short_circuit_in_if() { assert_eq!(run_c("int main() { int x=0; if (1 || (x=1)) {} printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_in_while() { assert_eq!(run_c("int main() { int x=0; while(0 && (x=1)) {} printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_increment() { assert_eq!(run_c("int main() { int x=0; 1 || x++; printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_ternary() { assert_eq!(run_c("int main() { int x=0; (1 || x++) ? (x+=10) : (x+=100); printf(\"%d\", x); return 0; }"), vec!["10"]); }
#[test] fn short_circuit_sequence_point() { assert_eq!(run_c("int main() { int x=1; x == 1 && (x=2); printf(\"%d\", x); return 0; }"), vec!["2"]); } // Safe, sequence point at &&
#[test] fn short_circuit_comma_operator() { assert_eq!(run_c("int main() { int x=0; 0 && (x=1, 1); printf(\"%d\", x); return 0; }"), vec!["0"]); }
#[test] fn short_circuit_array_bounds() { assert_eq!(run_c("int main() { int arr[2]={1,2}; int i=5; if (i<2 && arr[i]==1) printf(\"yes\"); else printf(\"no\"); return 0; }"), vec!["no"]); } // Protects bounds
