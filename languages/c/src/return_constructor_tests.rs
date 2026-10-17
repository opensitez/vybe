thread_local! {
    pub(super) static REFERENCE_RETURNS: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

#[test]
fn direct_return_constructor_preserves_complete_typed_ast_and_ordered_context() {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REFERENCE_RETURNS.with(|flag| flag.set(self.0)); }
    }
    let mut handoffs = 0;
    for source in [
        "int f(void){return 3;}",
        "int f(char *p,int i){return p[i++];}",
        "char f(char x){return x;}",
        "unsigned short f(unsigned short *p){return *p;}",
        "unsigned int f(unsigned int *p){return p[1];}",
        "long long f(long long *p){return p[1];}",
        "unsigned long long f(unsigned long long *p){return *p;}",
        "typedef unsigned long long W;W f(W *p){return p[1];}",
        "float f(float *p){return *p;}",
        "double f(double *p){return p[1];}",
        "int *f(int *p,int i){return &p[i++];}",
        "struct R{int x;};struct R *f(struct R *p){return p;}",
        "struct R{int x;};struct R f(struct R value){return value;}",
        "struct R{int x;};struct R get(void);struct R f(void){return get();}",
        "int f(int x){return (x++,observe(x),x*2);}",
        "int f(int x){return x && ready() ? selected(x) : other();}",
        "int f(int n){int values[n++];return sizeof(values)+n;}",
        "int f(int x){return ({int y=x;y+1;});}",
        "typedef int T;T f(T x){{typedef unsigned short T;T y=3;if(x)return y;}return x;}",
        "int f(int x){switch(x){case 1:return x+2;default:return x-1;}}",
        "int f(int x){while(x){if(x==2)return x;--x;}return 0;}",
        "int f(int x){if(x)goto done;return 2;done:return 3;}",
        "int f(int x){int inner(int y){return y+1;}return inner(x);}",
        "void f(int *p){if(p){*p=2;return;}return;}",
        "#define NEXT(x) ((x)+1)\nint f(int x){return NEXT(x);}",
    ] {
        let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
        for reference in [true, false] {
            let _restore = Restore(REFERENCE_RETURNS.with(|flag| flag.replace(reference)));
            let (module, counts) = crate::frontend::profile_construction(|| crate::parse_optimized(source));
            assert_eq!(format!("{:?}", module.unwrap()), expected, "reference={reference}: {source}");
            if reference { assert_eq!(counts.direct_return_handoffs, 0, "{source}"); }
            else {
                if !source.starts_with("void ") {
                    assert!(counts.direct_return_handoffs > 0, "value return did not consume its constructor: {source}");
                }
                handoffs += counts.direct_return_handoffs;
            }
        }
    }
    assert!(handoffs > 0, "actual Native ordered walk must consume owned value returns");
}

#[test]
fn void_returns_keep_complete_control_handoff_without_value_return_receipts() {
    for source in ["void f(void){return;}", "void f(int x){if(x)return;else return;}"] {
        let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
        let (module, counts) = crate::frontend::profile_construction(|| crate::parse_optimized(source));
        assert_eq!(format!("{:?}", module.unwrap()), expected, "{source}");
        assert_eq!(counts.direct_return_handoffs, 0, "void return is already complete: {source}");
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn direct_return_constructor_executes_identical_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/frontend_return_handoffs.c")]);
}
