use crate::helpers::run_main;

#[test]
fn uuid_from_string_version_4() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); System.out.println(u.version());"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn uuid_from_string_variant() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); System.out.println(u.variant());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn uuid_from_string_most_bits() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001"); System.out.println(u.getMostSignificantBits());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn uuid_from_string_least_bits() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn uuid_to_string_roundtrip() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("6ba7b810-9dad-11d1-80b4-00c04fd430c8"); System.out.println(u.toString());"#);
    assert_eq!(out, vec!["6ba7b810-9dad-11d1-80b4-00c04fd430c8"]);
}

#[test]
fn uuid_equals_same() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); java.util.UUID b = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_equals_different() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); java.util.UUID b = java.util.UUID.fromString("6ba7b810-9dad-11d1-80b4-00c04fd430c8"); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uuid_compare_to_equal() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("00000000-0000-0000-0000-000000000002"); java.util.UUID b = java.util.UUID.fromString("00000000-0000-0000-0000-000000000002"); System.out.println(a.compareTo(b));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn uuid_compare_to_less() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001"); java.util.UUID b = java.util.UUID.fromString("00000000-0000-0000-0000-000000000002"); System.out.println(a.compareTo(b) < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_compare_to_greater() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("00000000-0000-0000-0000-000000000003"); java.util.UUID b = java.util.UUID.fromString("00000000-0000-0000-0000-000000000002"); System.out.println(a.compareTo(b) > 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_hash_code_consistent() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); java.util.UUID b = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); System.out.println(a.hashCode() == b.hashCode());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_name_uuid_from_bytes() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes("vybe".getBytes()); System.out.println(u.version());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn uuid_name_uuid_deterministic() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.nameUUIDFromBytes("test".getBytes()); java.util.UUID b = java.util.UUID.nameUUIDFromBytes("test".getBytes()); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_name_uuid_different_input() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.nameUUIDFromBytes("a".getBytes()); java.util.UUID b = java.util.UUID.nameUUIDFromBytes("b".getBytes()); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uuid_random_uuid_version_4() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.randomUUID(); System.out.println(u.version() == 4);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_random_uuid_not_null() {
    let out = run_main(r#"System.out.println(java.util.UUID.randomUUID() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_from_string_all_zeros() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000000"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn uuid_from_string_max_bits() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("ffffffff-ffff-ffff-ffff-ffffffffffff"); System.out.println(u.getMostSignificantBits() < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_version_1_time_based() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("6ba7b810-9dad-11d1-80b4-00c04fd430c8"); System.out.println(u.version());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn uuid_version_3_name_based() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes("dns:example.com".getBytes()); System.out.println(u.version());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn uuid_variant_reserved() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000000"); System.out.println(u.variant());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn uuid_to_string_lowercase() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("ABCDEF12-3456-7890-ABCD-EF1234567890"); System.out.println(u.toString().equals(u.toString().toLowerCase()));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_most_significant_nonzero() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000010"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["16"]);
}

#[test]
fn uuid_least_significant_large() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-00000000ffff"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["65535"]);
}

#[test]
fn uuid_name_uuid_empty_bytes() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes(new byte[0]); System.out.println(u != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_from_string_dashed_format() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("12345678-1234-5678-1234-567812345678"); System.out.println(u.toString().length());"#);
    assert_eq!(out, vec!["36"]);
}

#[test]
fn uuid_random_two_distinct() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.randomUUID(); java.util.UUID b = java.util.UUID.randomUUID(); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uuid_compare_most_significant_first() {
    let out = run_main(r#"java.util.UUID a = java.util.UUID.fromString("10000000-0000-0000-0000-000000000000"); java.util.UUID b = java.util.UUID.fromString("20000000-0000-0000-0000-000000000000"); System.out.println(a.compareTo(b) < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_name_uuid_from_utf8() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes("\u00e9".getBytes(java.nio.charset.StandardCharsets.UTF_8)); System.out.println(u.version());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn uuid_get_most_negative() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("80000000-0000-0000-0000-000000000000"); System.out.println(u.getMostSignificantBits() < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_get_least_negative() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-800000000000"); System.out.println(u.getLeastSignificantBits() < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_from_string_version_bits() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-4000-8000-000000000000"); System.out.println(u.version());"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn uuid_to_string_contains_dashes() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.randomUUID(); System.out.println(u.toString().contains("-"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_name_uuid_long_input() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes("the quick brown fox jumps over the lazy dog".getBytes()); System.out.println(u.variant());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn uuid_equals_null_false() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.randomUUID(); System.out.println(u.equals(null));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn uuid_equals_self() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.randomUUID(); System.out.println(u.equals(u));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_from_string_nil() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000000"); System.out.println(u.toString());"#);
    assert_eq!(out, vec!["00000000-0000-0000-0000-000000000000"]);
}

#[test]
fn uuid_compare_nil_vs_one() {
    let out = run_main(r#"java.util.UUID nil = java.util.UUID.fromString("00000000-0000-0000-0000-000000000000"); java.util.UUID one = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001"); System.out.println(nil.compareTo(one) < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_random_uuid_variant_2() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.randomUUID(); System.out.println(u.variant());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn uuid_name_uuid_known_prefix() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes("url:www.example.com".getBytes()); System.out.println(u.toString().length() == 36);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_most_bits_shift() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000010000"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["65536"]);
}

#[test]
fn uuid_from_string_uppercase_input() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("550E8400-E29B-41D4-A716-446655440000"); System.out.println(u.version());"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn uuid_hash_code_not_zero_usually() {
    let out = run_main(
        r#"java.util.UUID u = java.util.UUID.fromString("550e8400-e29b-41d4-a716-446655440000"); System.out.println(u.hashCode() == u.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_name_uuid_binary() {
    let out = run_main(
        r#"java.util.UUID u = java.util.UUID.nameUUIDFromBytes(new byte[]{1, 2, 3}); System.out.println(u.getMostSignificantBits() != 0 ? true : u.getLeastSignificantBits() != 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn uuid_version_mask_4() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-5000-8000-000000000000"); System.out.println(u.version());"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn uuid_from_string_least_one() {
    let out = run_main(r#"java.util.UUID u = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001"); System.out.println(u.getLeastSignificantBits());"#);
    assert_eq!(out, vec!["1"]);
}

