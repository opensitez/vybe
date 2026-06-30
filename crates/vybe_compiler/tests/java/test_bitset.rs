use crate::helpers::run_main;

#[test]
fn bitset_new_is_empty() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); System.out.println(bs.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_set_makes_nonempty() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); System.out.println(bs.isEmpty());"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_get_unset_returns_false() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); System.out.println(bs.get(5));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_get_set_returns_true() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(5); System.out.println(bs.get(5));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_clear_unsets_bit() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(3); bs.clear(3); System.out.println(bs.get(3));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_flip_toggles_bit() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.flip(2); System.out.println(bs.get(2));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_flip_twice_restores() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.flip(2); bs.flip(2); System.out.println(bs.get(2));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_cardinality_counts_set_bits() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(1); bs.set(3); System.out.println(bs.cardinality());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn bitset_length_upper_bound() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(10); System.out.println(bs.length());"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn bitset_size_includes_highest() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(7); System.out.println(bs.size() >= 8);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_next_set_bit_first() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(4); System.out.println(bs.nextSetBit(0));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn bitset_next_set_bit_from_index() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(2); bs.set(8); System.out.println(bs.nextSetBit(3));"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitset_next_clear_bit() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); System.out.println(bs.nextClearBit(0));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bitset_previous_set_bit() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(5); System.out.println(bs.previousSetBit(10));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn bitset_previous_clear_bit() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(5); System.out.println(bs.previousClearBit(5));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn bitset_and_intersection() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(1); a.set(2); java.util.BitSet b = new java.util.BitSet(); b.set(2); b.set(3); a.and(b); System.out.println(a.cardinality());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bitset_or_union() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(0); java.util.BitSet b = new java.util.BitSet(); b.set(1); a.or(b); System.out.println(a.cardinality());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn bitset_xor_symmetric_diff() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(0); a.set(1); java.util.BitSet b = new java.util.BitSet(); b.set(1); a.xor(b); System.out.println(a.cardinality());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bitset_and_not_difference() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(0); a.set(1); java.util.BitSet b = new java.util.BitSet(); b.set(1); a.andNot(b); System.out.println(a.get(0)); System.out.println(a.get(1));"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn bitset_intersects_true() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(3); java.util.BitSet b = new java.util.BitSet(); b.set(3); System.out.println(a.intersects(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_intersects_false() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(1); java.util.BitSet b = new java.util.BitSet(); b.set(2); System.out.println(a.intersects(b));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_set_range() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(2, 5); System.out.println(bs.cardinality());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn bitset_clear_range() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0, 10); bs.clear(3, 7); System.out.println(bs.cardinality());"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitset_flip_range() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0, 4); bs.flip(1, 3); System.out.println(bs.get(1)); System.out.println(bs.get(2));"#);
    assert_eq!(out, vec!["false", "false"]);
}

#[test]
fn bitset_get_range_count() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); bs.set(2); System.out.println(bs.get(0, 3).cardinality());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn bitset_equals_same_bits() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(1); java.util.BitSet b = new java.util.BitSet(); b.set(1); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_equals_different() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(1); java.util.BitSet b = new java.util.BitSet(); b.set(2); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_hash_code_equal_sets() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(4); java.util.BitSet b = new java.util.BitSet(); b.set(4); System.out.println(a.hashCode() == b.hashCode());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_to_string_format() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); bs.set(2); System.out.println(bs.toString().contains("0"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_clone_copies_bits() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(6); java.util.BitSet b = (java.util.BitSet) a.clone(); System.out.println(b.get(6));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_set_with_value_true() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(3, true); System.out.println(bs.get(3));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_set_with_value_false() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(3); bs.set(3, false); System.out.println(bs.get(3));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn bitset_next_set_bit_minus_one() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); System.out.println(bs.nextSetBit(0));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bitset_stream_count() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(1); bs.set(4); System.out.println(bs.stream().count());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn bitset_cardinality_after_clear_all() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); bs.clear(); System.out.println(bs.cardinality());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bitset_high_index() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(63); System.out.println(bs.get(63));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_or_into_self() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(1); a.or(a); System.out.println(a.cardinality());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn bitset_and_with_empty() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(5); java.util.BitSet b = new java.util.BitSet(); a.and(b); System.out.println(a.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_xor_self_clears() {
    let out = run_main(r#"java.util.BitSet a = new java.util.BitSet(); a.set(2); a.xor(a); System.out.println(a.isEmpty());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_value_of_long() {
    let out = run_main(r#"java.util.BitSet bs = java.util.BitSet.valueOf(5L); System.out.println(bs.get(0)); System.out.println(bs.get(2));"#);
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn bitset_to_long_array() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); System.out.println(bs.toLongArray().length >= 1);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_value_of_byte_array() {
    let out = run_main(r#"java.util.BitSet bs = java.util.BitSet.valueOf(new byte[]{1}); System.out.println(bs.get(0));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_to_byte_array() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(0); System.out.println(bs.toByteArray().length >= 1);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn bitset_previous_set_bit_none() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); System.out.println(bs.previousSetBit(5));"#);
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bitset_set_high_then_length() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(100); System.out.println(bs.length());"#);
    assert_eq!(out, vec!["101"]);
}

#[test]
fn bitset_clear_single_keeps_others() {
    let out = run_main(r#"java.util.BitSet bs = new java.util.BitSet(); bs.set(1); bs.set(3); bs.clear(1); System.out.println(bs.get(3));"#);
    assert_eq!(out, vec!["true"]);
}

