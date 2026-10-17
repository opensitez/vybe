use super::*;
use std::cell::Cell;

thread_local! {
    pub(super) static REFERENCE_RESOLUTION: Cell<bool> = const { Cell::new(false) };
}

pub(super) fn uncached(walker: &Walker, text: &str) -> String {
    let mut out = text.trim().to_string();
    for (name, target) in &walker.typedef_scalar_alias_order {
        out = replace_word(&out, name, target);
    }
    out
}

fn with_reference<T>(enabled: bool, action: impl FnOnce() -> T) -> T {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) {
            REFERENCE_RESOLUTION.with(|flag| flag.set(self.0));
        }
    }
    let _restore = Restore(REFERENCE_RESOLUTION.with(|flag| flag.replace(enabled)));
    action()
}

#[test]
fn memo_preserves_order_boundaries_quotes_and_exact_spelling() {
    let mut walker = Walker::default();
    for (name, target) in [
        ("Word", "Leaf"),
        ("Leaf", "unsigned short"),
        ("WordLong", "long long"),
        ("Other", "Word"),
        ("X", "long"),
    ] {
        walker.insert_typedef_scalar_alias(name.into(), target.into());
    }
    for source in [
        "Word",
        "  const Word * volatile [3]  ",
        "WordLong Word",
        "Other",
        "WordSuffix PrefixWord Word_ Word2",
        "\"Word\" Word 'X' X",
        "\"escaped \\\" Word\" Word",
        "Word\t*\nLeaf",
        "\u{3bb} Word",
        "",
        "  ",
    ] {
        let expected = uncached(&walker, source);
        assert_eq!(
            walker.resolve_typedef_scalar_aliases(source),
            expected,
            "{source:?}"
        );
        assert_eq!(
            walker.resolve_typedef_scalar_aliases(source),
            expected,
            "cache hit: {source:?}"
        );
    }
    assert!(walker.typedef_scalar_memo.borrow().hits > 0);
}

#[test]
fn repeated_type_queries_skip_all_replacement_passes_on_hits() {
    let mut walker = Walker::default();
    for index in 0..32 {
        walker.insert_typedef_scalar_alias(format!("Alias{index:02}"), "unsigned int".into());
    }
    let expected = uncached(&walker, "const Alias31 *");
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const Alias31 *"),
        expected
    );
    for _ in 0..400 {
        assert_eq!(
            walker.resolve_typedef_scalar_aliases("  const Alias31 *  "),
            expected
        );
    }
    let memo = walker.typedef_scalar_memo.borrow();
    assert_eq!(memo.misses, 1);
    assert_eq!(memo.hits, 400);
    assert_eq!(memo.replacement_passes, 32);
}

#[test]
fn insertion_redefinition_removal_and_nested_restoration_invalidate_all_entries() {
    let mut walker = Walker::default();
    walker.insert_typedef_scalar_alias("Unrelated".into(), "int".into());
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "T");
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "T");
    walker.insert_typedef_scalar_alias("T".into(), "int".into());
    assert!(walker.typedef_scalar_memo.borrow().values.is_empty());
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "int");
    let outer = walker.typedef_scalar_aliases["T"].clone();
    walker.insert_typedef_scalar_alias("T".into(), "double".into());
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "double");
    walker.insert_typedef_scalar_alias("T".into(), "unsigned short".into());
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const T *"),
        "const unsigned short *"
    );
    walker.insert_typedef_scalar_alias("T".into(), "double".into());
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const T *"),
        "const double *"
    );
    walker.insert_typedef_scalar_alias("T".into(), outer);
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "int");
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const T *"),
        "const int *"
    );
    walker.remove_typedef_scalar_alias("T");
    assert!(walker.typedef_scalar_memo.borrow().values.is_empty());
    assert_eq!(walker.resolve_typedef_scalar_aliases("T"), "T");
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const T *"),
        "const T *"
    );
    assert_eq!(walker.resolve_typedef_scalar_aliases("Unrelated"), "int");
}

#[test]
fn alias_target_changes_invalidate_dependent_cached_spellings() {
    let mut walker = Walker::default();
    walker.insert_typedef_scalar_alias("Outer".into(), "Inner".into());
    walker.insert_typedef_scalar_alias("Inner".into(), "int".into());
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const Outer *"),
        "const int *"
    );
    walker.insert_typedef_scalar_alias("Inner".into(), "double".into());
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const Outer *"),
        "const double *"
    );
    walker.remove_typedef_scalar_alias("Inner");
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("const Outer *"),
        "const Inner *"
    );
}

#[test]
fn memo_retention_is_bounded_by_entry_count_and_text_bytes() {
    let mut walker = Walker::default();
    walker.insert_typedef_scalar_alias("T".into(), "int".into());
    for index in 0..TYPEDEF_SCALAR_MEMO_ENTRIES * 4 {
        let text = format!("const Unknown{index} *");
        assert_eq!(
            walker.resolve_typedef_scalar_aliases(&text),
            uncached(&walker, &text)
        );
        let memo = walker.typedef_scalar_memo.borrow();
        assert!(memo.values.len() <= TYPEDEF_SCALAR_MEMO_ENTRIES);
        assert!(memo.bytes <= TYPEDEF_SCALAR_MEMO_BYTES);
        assert_eq!(
            memo.bytes,
            memo.values
                .iter()
                .map(|(key, value)| key.len() + value.len())
                .sum::<usize>()
        );
    }
    for index in 0..64 {
        let text = format!("{}_{index}", "q".repeat(2048));
        assert_eq!(
            walker.resolve_typedef_scalar_aliases(&text),
            uncached(&walker, &text)
        );
        let memo = walker.typedef_scalar_memo.borrow();
        assert!(memo.values.len() <= TYPEDEF_SCALAR_MEMO_ENTRIES);
        assert!(memo.bytes <= TYPEDEF_SCALAR_MEMO_BYTES);
    }
    let oversized = "q".repeat(TYPEDEF_SCALAR_MEMO_BYTES + 1);
    assert_eq!(walker.resolve_typedef_scalar_aliases(&oversized), oversized);
    assert!(
        !walker
            .typedef_scalar_memo
            .borrow()
            .values
            .contains_key(&oversized)
    );
    walker.insert_typedef_scalar_alias("T".into(), "q".repeat(TYPEDEF_SCALAR_MEMO_BYTES + 1));
    assert_eq!(
        walker.resolve_typedef_scalar_aliases("T").len(),
        TYPEDEF_SCALAR_MEMO_BYTES + 1
    );
    let memo = walker.typedef_scalar_memo.borrow();
    assert!(memo.values.is_empty());
    assert_eq!(memo.bytes, 0);
}

#[test]
fn empty_alias_table_avoids_memoization_and_walkers_do_not_share_entries() {
    let first = Walker::default();
    assert_eq!(first.resolve_typedef_scalar_aliases("  int *  "), "int *");
    assert_eq!(first.resolve_typedef_scalar_aliases("int *"), "int *");
    assert!(first.typedef_scalar_memo.borrow().values.is_empty());
    assert_eq!(first.typedef_scalar_memo.borrow().hits, 0);
    assert_eq!(first.typedef_scalar_memo.borrow().replacement_passes, 0);
    let mut second = Walker::default();
    second.insert_typedef_scalar_alias("int".into(), "double".into());
    assert_eq!(second.resolve_typedef_scalar_aliases("int *"), "double *");
    assert_eq!(first.resolve_typedef_scalar_aliases("int *"), "int *");
}

#[test]
fn memo_and_reference_resolution_preserve_complete_peg_and_native_modules() {
    for source in [
        "typedef unsigned short Half;typedef Half Word;static Word values[]={1,2,3}; \
         Word total(int n){Word sum=0;for(int i=0;i<n;i++)sum+=values[i];return sum;} \
         int main(void){return total(3)!=6;}",
        "typedef int T;int f(void){T a=1;{typedef unsigned short T;T b=2;a+=b;} \
         T c=3;return a+c;}int main(void){return f();}",
        "typedef int T;int f(void){typedef struct{int x;}T;T a={3};return a.x;}",
        "typedef unsigned int Word;typedef Word *Pointer; \
         Word f(Pointer p){return p[0]+p[1];}int main(void){Word a[]={1,2};return f(a)!=3;}",
        "typedef int (*Callback)(int,...);int value(int x,...){return x;} \
         static Callback table[2]={value,value};int main(void){return table[1](7,0)!=7;}",
    ] {
        for parse in [crate::parse_with_peg, crate::parse_optimized] {
            let expected = with_reference(true, || parse(source)).map(|value| format!("{value:?}"));
            let actual = with_reference(false, || parse(source)).map(|value| format!("{value:?}"));
            assert_eq!(actual, expected, "{source}");
            assert!(actual.is_ok(), "{actual:?}: {source}");
        }
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

fn parse_cached(source: &str) -> Result<vybe_ast::Module, String> {
    with_reference(false, || crate::parse_optimized(source))
}

fn parse_reference(source: &str) -> Result<vybe_ast::Module, String> {
    with_reference(true, || crate::parse_optimized(source))
}

#[test]
fn memo_preserves_standard_wasm_owned_emission_vm_and_noeq() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(
        crate::profile_source(),
        [parse_reference, parse_cached],
        [
            "typedef unsigned short Half;typedef Half Word;static Word values[]={1,2,3}; \
             Word total(int n){Word sum=0;for(int i=0;i<n;i++)sum+=values[i];return sum;} \
             int main(void){return total(3)!=6;}",
            "typedef int (*Callback)(int,...);int value(int x,...){return x;} \
             static Callback table[2]={value,value};int main(void){return table[1](7,0)!=7;}",
        ],
    );
}
