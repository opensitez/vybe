use crate::helpers::run_in_main;

const TYPES: &str = r#"
enum Color { RED, GREEN, BLUE, YELLOW }
"#;

#[test]
fn enum_set_none_of_is_empty() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.noneOf(Color.class); System.out.println(s.isEmpty());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_all_of_contains_all() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.allOf(Color.class); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn enum_set_of_single() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_of_pair() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED, Color.BLUE); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_of_triple() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED, Color.GREEN, Color.BLUE); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enum_set_add_increases_size() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.noneOf(Color.class); s.add(Color.GREEN); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_remove_decreases_size() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); s.remove(Color.RED); System.out.println(s.isEmpty());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_contains_present() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.BLUE); System.out.println(s.contains(Color.BLUE));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_contains_absent() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.contains(Color.GREEN));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_add_all_union() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.RED); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.BLUE); a.addAll(b); System.out.println(a.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_remove_all() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.allOf(Color.class); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.RED, Color.GREEN); a.removeAll(b); System.out.println(a.contains(Color.RED));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_retain_all() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.allOf(Color.class); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.RED, Color.BLUE); a.retainAll(b); System.out.println(a.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_complement_of() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.complementOf(java.util.EnumSet.of(Color.RED)); System.out.println(s.contains(Color.GREEN)); System.out.println(s.contains(Color.RED));"#, TYPES);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn enum_set_copy_of() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.YELLOW); java.util.EnumSet<Color> b = java.util.EnumSet.copyOf(a); System.out.println(b.contains(Color.YELLOW));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_range_inclusive() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.range(Color.RED, Color.BLUE); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enum_set_iterator_next() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.iterator().next());"#, TYPES);
    assert_eq!(out, vec!["RED"]);
}

#[test]
fn enum_set_equals_same() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.GREEN); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.GREEN); System.out.println(a.equals(b));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_equals_different() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.GREEN); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.BLUE); System.out.println(a.equals(b));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_hash_code_equal() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.RED); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.RED); System.out.println(a.hashCode() == b.hashCode());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_to_string_contains() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.toString().contains("RED"));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_clear_empties() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.allOf(Color.class); s.clear(); System.out.println(s.isEmpty());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_add_duplicate_no_growth() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); s.add(Color.RED); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_remove_absent_no_op() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); s.remove(Color.BLUE); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_contains_all_true() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.allOf(Color.class); java.util.EnumSet<Color> b = java.util.EnumSet.of(Color.RED); System.out.println(a.containsAll(b));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_contains_all_false() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.RED); java.util.EnumSet<Color> b = java.util.EnumSet.allOf(Color.class); System.out.println(a.containsAll(b));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_of_four() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED, Color.GREEN, Color.BLUE, Color.YELLOW); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn enum_set_none_of_size_zero() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.noneOf(Color.class); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn enum_set_all_of_contains_red() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.allOf(Color.class); System.out.println(s.contains(Color.RED));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_complement_size() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.complementOf(java.util.EnumSet.of(Color.RED, Color.GREEN)); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_range_single() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.range(Color.BLUE, Color.BLUE); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_copy_independent() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.RED); java.util.EnumSet<Color> b = java.util.EnumSet.copyOf(a); b.add(Color.BLUE); System.out.println(a.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_add_all_self() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); s.addAll(s); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_set_remove_all_self_clears() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); s.removeAll(s); System.out.println(s.isEmpty());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_retain_all_self() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED, Color.GREEN); s.retainAll(s); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_iterator_has_next() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.iterator().hasNext());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_of_green() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.GREEN); System.out.println(s.iterator().next());"#, TYPES);
    assert_eq!(out, vec!["GREEN"]);
}

#[test]
fn enum_set_range_red_to_yellow() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.range(Color.RED, Color.YELLOW); System.out.println(s.size());"#, TYPES);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn enum_set_complement_of_all() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.complementOf(java.util.EnumSet.allOf(Color.class)); System.out.println(s.isEmpty());"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_add_returns_true_new() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.noneOf(Color.class); System.out.println(s.add(Color.BLUE));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_add_returns_false_dup() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.BLUE); System.out.println(s.add(Color.BLUE));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_remove_returns_true() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.remove(Color.RED));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_remove_returns_false() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.of(Color.RED); System.out.println(s.remove(Color.GREEN));"#, TYPES);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn enum_set_none_of_class() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.noneOf(Color.class); System.out.println(s.getClass().getSimpleName().contains("EnumSet"));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_all_of_class() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.allOf(Color.class); System.out.println(s.contains(Color.YELLOW));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_set_copy_of_from_collection() {
    let out = run_in_main(r#"java.util.EnumSet<Color> a = java.util.EnumSet.of(Color.RED, Color.BLUE); java.util.EnumSet<Color> b = java.util.EnumSet.copyOf(a); System.out.println(b.size());"#, TYPES);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_set_complement_preserves_type() {
    let out = run_in_main(r#"java.util.EnumSet<Color> s = java.util.EnumSet.complementOf(java.util.EnumSet.of(Color.YELLOW)); System.out.println(s.contains(Color.RED));"#, TYPES);
    assert_eq!(out, vec!["true"]);
}

