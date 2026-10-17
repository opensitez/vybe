thread_local! {
    pub(super) static REFERENCE_FUNCTIONS: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

#[test]
fn direct_function_constructor_preserves_typed_signatures_scopes_and_normalization() {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REFERENCE_FUNCTIONS.with(|flag| flag.set(self.0)); }
    }
    let mut handoffs = 0;
    for source in [
        "int f(void){return 3;}",
        "typedef unsigned short W;double f(int a,W b,double c){return a+b+c;}",
        "int *f(int *p){return p;}",
        "struct S{int x;};struct S f(struct S s){return s;}",
        "typedef int (*F)(int);int f(F p,int x){return p(x);}",
        "int f(int a[3]){return a[2];}",
        "int f(int n,int a[n]){return a[n-1];}",
        "int f(int n,int (*a)[3]){return a[n][2];}",
        "int f(int x,...){return x;}",
        "static inline int f(int n){if(n)return f(n-1);return 0;}",
        "int main(int argc,char **argv){return argc>0 && argv!=0;}",
        "typedef int T;T f(T x){T y=x;{T x=2;y+=x;}return y;}",
        "int f(int n){int a[n++];a[0]=n;return a[0];}",
        "void f(int *p){if(p){*p=2;return;}return;}",
        "unsigned f(unsigned n){switch(n){case 0:return 1;default:return n+1;}}",
        "int f(int x){int y=2;return x+y;}int g(int y){return f(y);}",
        "#define STEP(x) ((x)+1)\nint f(int x){return STEP(x);}",
        "int (f)(int x){return x;}",
    ] {
        let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
        for reference in [true, false] {
            let _restore = Restore(REFERENCE_FUNCTIONS.with(|flag| flag.replace(reference)));
            let (module, counts) = crate::frontend::profile_construction(|| crate::parse_optimized(source));
            assert_eq!(format!("{:?}", module.unwrap()), expected, "reference={reference}: {source}");
            if reference { assert_eq!(counts.direct_function_handoffs, 0); }
            else { handoffs += counts.direct_function_handoffs; }
        }
    }
    assert!(handoffs > 0, "actual Native ordered walk must consume owned functions");
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn direct_function_constructor_executes_identical_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/frontend_function_handoffs.c")]);
}
