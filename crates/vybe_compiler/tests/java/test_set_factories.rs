/// Set factory methods: Set.of, Set.copyOf, immutability.
use crate::helpers::run_main;

#[test]
fn set_of_distinct_elements_supports_contains() {
    let out = run_main(
        r#"java.util.Set<Integer> s = java.util.Set.of(1, 2, 3); System.out.println(s.contains(2)); System.out.println(s.size());"#,
    );
    assert_eq!(out, vec!["true", "3"]);
}

#[test]
fn set_of_empty_is_empty() {
    let out =
        run_main("java.util.Set<String> s = java.util.Set.of(); System.out.println(s.isEmpty());");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn set_of_add_throws_unsupported_operation() {
    let out = run_main(
        r#"java.util.Set<Integer> s = java.util.Set.of(1); try { s.add(2); System.out.println("mutated"); } catch (UnsupportedOperationException e) { System.out.println("immutable"); }"#,
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn set_copy_of_reflects_source_elements() {
    let out = run_main(
        r#"java.util.Set<String> copy = java.util.Set.copyOf(java.util.Arrays.asList("a", "b")); System.out.println(copy.contains("b")); System.out.println(copy.size());"#,
    );
    assert_eq!(out, vec!["true", "2"]);
}

#[test]
fn set_of_duplicate_argument_throws_illegal_argument() {
    let out = run_main(
        r#"try { java.util.Set.of(1, 1); System.out.println("ok"); } catch (IllegalArgumentException e) { System.out.println("dup"); }"#,
    );
    assert_eq!(out, vec!["dup"]);
}

#[test]
fn set_of_remove_throws_unsupported_operation() {
    let out = run_main(
        r#"java.util.Set<String> s = java.util.Set.of("x"); try { s.remove("x"); System.out.println("mutated"); } catch (UnsupportedOperationException e) { System.out.println("immutable"); }"#,
    );
    assert_eq!(out, vec!["immutable"]);
}
