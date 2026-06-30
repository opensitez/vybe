/// Map factory methods: Map.of, Map.ofEntries, immutability.
use crate::helpers::run_main;

#[test]
fn map_of_two_entries_supports_get() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.of("a", 1, "b", 2); System.out.println(m.get("a")); System.out.println(m.get("b"));"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn map_of_empty_has_size_zero() {
    let out = run_main(
        "java.util.Map<String, Integer> m = java.util.Map.of(); System.out.println(m.isEmpty());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn map_of_single_entry_contains_key() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.of("solo", 9); System.out.println(m.containsKey("solo")); System.out.println(m.size());"#,
    );
    assert_eq!(out, vec!["true", "1"]);
}

#[test]
fn map_of_rejects_null_key_at_runtime() {
    let out = run_main(
        r#"try { java.util.Map.of(null, 1); System.out.println("ok"); } catch (NullPointerException e) { System.out.println("npe"); }"#,
    );
    assert_eq!(out, vec!["npe"]);
}

#[test]
fn map_of_put_throws_unsupported_operation() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.of("k", 1); try { m.put("x", 2); System.out.println("mutated"); } catch (UnsupportedOperationException e) { System.out.println("immutable"); }"#,
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn map_of_entries_builds_from_entry_objects() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.ofEntries(java.util.Map.entry("x", 10), java.util.Map.entry("y", 20)); System.out.println(m.get("y"));"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn map_of_get_or_default_on_missing_key() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.of("a", 1); System.out.println(m.getOrDefault("z", 99));"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn map_of_key_set_contains_declared_keys() {
    let out = run_main(
        r#"java.util.Map<String, Integer> m = java.util.Map.of("p", 1, "q", 2); System.out.println(m.keySet().contains("q"));"#,
    );
    assert_eq!(out, vec!["true"]);
}
