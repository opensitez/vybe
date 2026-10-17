thread_local! {
    pub(super) static REFERENCE_PLACES: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

#[test]
fn direct_index_member_constructors_preserve_complete_typed_ast_and_context() {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REFERENCE_PLACES.with(|flag| flag.set(self.0)); }
    }
    let mut indices = 0;
    let mut members = 0;
    for source in [
        "int f(){int a[3]={1,2,3};int i=0;a[i++]+=2;return a[0]+i;}",
        "struct S{int x;};int f(){struct S s;s.x=2;return s.x++;}",
        "struct S{int x;};int f(struct S *p,int i){p[i].x++;return p[i].x;}",
        "struct S{int *p;};int f(struct S *s,int i){s->p[i]+=2;return s->p[i];}",
        "struct S{int *p;};int *f(struct S *s,int i){return &s[1].p[i];}",
        "int f(char *s){int i=0;return s[i++]+i;}",
        "unsigned short f(unsigned short *p,int i){return p[i];}",
        "int f(int (*p)[3]){p[1][2]=9;return p[1][2];}",
        "typedef int *P;int f(P p){return 2[p];}",
        "struct I{int x;};struct O{struct I i;};int f(struct O *p){return p->i.x;}",
        "struct S{int x;};int f(struct S *a,struct S *b,int c){return (c?a:b)->x;}",
        "union U{int x;unsigned y;};int f(union U *p){p->x=3;return p->y;}",
        "int f(int n){int a[n];a[n-1]=3;int *p=&a[n-1];return *p;}",
        "struct S{int x;};struct S *owner(void);int index(void);int f(){return owner()[index()].x++;}",
        "struct S{int x;int (*f)(int);};int g(struct S *p){return p->f(p->x);}",
        "int f(){return \"abc\"[1];}",
    ] {
        let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
        for reference in [true, false] {
            let _restore = Restore(REFERENCE_PLACES.with(|flag| flag.replace(reference)));
            let (module, counts) = crate::frontend::profile_construction(|| crate::parse_optimized(source));
            assert_eq!(format!("{:?}", module.unwrap()), expected, "reference={reference}: {source}");
            if reference {
                assert_eq!(counts.direct_index_handoffs + counts.direct_member_handoffs, 0);
            } else {
                indices += counts.direct_index_handoffs;
                members += counts.direct_member_handoffs;
            }
        }
    }
    assert!(indices > 0, "ordinary source indexes must consume owned ASTs");
    assert!(members > 0, "ordinary source members must consume owned ASTs");
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn direct_index_member_constructors_execute_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/frontend_index_member_handoffs.c")]);
}

#[test]
fn indexed_pointer_member_address_updates_original_backing() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        ["typedef struct{int value;int *items;}R;static R records[2];static int values[3]={3,5,7};\
          int main(){records[1].items=values;int *p=&records[1].items[2];*p=13;return values[2]!=13;}"]);
}

#[test]
fn indexed_pointer_member_address_preserves_single_evaluation() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/indexed_pointer_member_address_effects.c")]);
}

#[test]
fn indexed_pointer_member_address_preserves_typed_widths_and_backing() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized],
        [include_str!("../../../tests/c/doom_startup/indexed_pointer_member_address_widths.c")]);
}
