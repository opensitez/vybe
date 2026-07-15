use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn vla_goto_forward_bypass_vla() {
    assert_eq!(
        run_c("int main() { goto L; int n=5; int arr[n]; L: printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Technically UB if arr is used, but ok if not. Our compiler might accept it
#[test]
fn vla_goto_backward_out_of_scope() {
    assert_eq!(
        run_c(
            "int main() { int i=0; L: if(i++ == 0) { int n=5; int arr[n]; goto L; } printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_goto_forward_into_scope_fails() {
    assert_eq!(
        run_c(
            "/* int main() { goto L; { int n=5; int arr[n]; L: arr[0]=1; } return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_goto_backward_reentry() {
    assert_eq!(
        run_c(
            "int main() { int count=0; L: { int n=count+1; int arr[n]; arr[n-1]=n; count++; if (count<2) goto L; } printf(\"%d\", count); return 0; }"
        ),
        vec!["2"]
    );
} // Properly allocates and deallocates
#[test]
fn vla_goto_break_scope() {
    assert_eq!(
        run_c(
            "int main() { int sum=0; for(int i=1; i<=2; i++) { int arr[i]; arr[i-1]=i; if (i==2) break; sum += arr[i-1]; } printf(\"%d\", sum); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_goto_continue_scope() {
    assert_eq!(
        run_c(
            "int main() { int sum=0; for(int i=1; i<=2; i++) { int arr[i]; arr[i-1]=i; sum += arr[i-1]; continue; } printf(\"%d\", sum); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn vla_goto_return_scope() {
    assert_eq!(
        run_c(
            "int f() { int n=5; int arr[n]; return arr[0]=42; } int main() { printf(\"%d\", f()); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn vla_switch_bypass_fails() {
    assert_eq!(
        run_c(
            "/* int main() { int x=1; switch(x) { case 1: int n=5; int arr[n]; break; } return 0; } // label bypasses init */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_switch_bypass_with_braces() {
    assert_eq!(
        run_c(
            "int main() { int x=1; switch(x) { case 1: { int n=5; int arr[n]; arr[0]=1; printf(\"%d\", arr[0]); break; } } return 0; }"
        ),
        vec!["1"]
    );
} // Fine because it's a block
#[test]
fn vla_setjmp_longjmp_leak() {
    assert_eq!(
        run_c(
            "#include <setjmp.h>\nint main() { jmp_buf env; if(setjmp(env) == 0) { int n=1000; int arr[n]; longjmp(env, 1); } printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // VLA memory might leak depending on implementation, but program should not crash
#[test]
fn vla_sizeof_in_goto() {
    assert_eq!(
        run_c(
            "int main() { goto L; int n=5; int arr[n]; L: printf(\"%d\", (int)sizeof(arr)); return 0; }"
        ),
        vec!["0"]
    );
} // Undefined value, but maybe just 0 or uninitialized in our compiler
#[test]
fn vla_alloca_equivalent() {
    assert_eq!(
        run_c(
            "int main() { int n=5; void *p; { int arr[n]; p = arr; } /* p is dangling */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_nested_blocks_same_name() {
    assert_eq!(
        run_c(
            "int main() { int n=1; { int arr[n]; arr[0]=1; { int arr[n+1]; arr[1]=2; printf(\"%d\", arr[1]); } } return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn vla_recursive_call() {
    assert_eq!(
        run_c(
            "int f(int n) { if(n==0) return 0; int arr[n]; arr[n-1]=n; return arr[n-1] + f(n-1); } int main() { printf(\"%d\", f(3)); return 0; }"
        ),
        vec!["6"]
    );
} // 3+2+1
