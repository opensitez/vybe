use crate::helpers::run_prints;

#[test]
fn test_uuid_from_string_and_to_string_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "123e4567-e89b-12d3-a456-426614174000"
            val id = java.util.UUID.fromString(source)
            println(id.toString())
            println(id.variant())
            println(id.version())
        }
    "#,
    );
    assert_eq!(out, &["123e4567-e89b-12d3-a456-426614174000", "2", "1"]);
}

#[test]
fn test_uuid_comparison_and_hash_stability() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            val b = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            println(a == b)
            println(a.hashCode() == b.hashCode())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_uuid_variant_and_version_bits() {
    let out = run_prints(
        r#"
        fun main() {
            val id = java.util.UUID.fromString("ffffffff-ffff-4fff-8fff-ffffffffffff")
            println(id.version())
            println(id.variant())
            println(id.variant() == 2)
        }
    "#,
    );
    assert_eq!(out, &["4", "2", "true"]);
}

#[test]
fn test_uuid_most_and_least_bits_access() {
    let out = run_prints(
        r#"
        fun main() {
            val id = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001")
            println(id.mostSignificantBits)
            println(id.leastSignificantBits)
            println(id.toString())
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "00000000-0000-0000-0000-000000000001"]);
}

#[test]
fn test_uuid_name_from_bytes_is_deterministic() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = "kotlin".toByteArray()
            val a = java.util.UUID.nameUUIDFromBytes(bytes)
            val b = java.util.UUID.nameUUIDFromBytes(bytes)
            println(a)
            println(a == b)
            println(a.version())
        }
    "#,
    );
    assert_eq!(out, &["7f1d0d8e-2f1f-3138-92e7-4b3d8f5ef2d6", "true", "3"]);
}

#[test]
fn test_uuid_random_uuid_format_shape() {
    let out = run_prints(
        r#"
        fun main() {
            val id = java.util.UUID.randomUUID()
            val text = id.toString()
            println(text.length)
            println(text[8] == '-')
            println(text[13] == '-')
            println(text[18] == '-')
            println(text[23] == '-')
        }
    "#,
    );
    assert_eq!(out, &["36", "true", "true", "true", "true"]);
}

#[test]
fn test_uuid_parse_invalid_throws() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                java.util.UUID.fromString("not-a-uuid")
                println("bad")
            } catch (e: IllegalArgumentException) {
                println(e::class.simpleName)
            }
        }
    "#,
    );
    assert_eq!(out, &["IllegalArgumentException"]);
}

#[test]
fn test_uuid_timestamped_time_ordering_for_v1_or_not() {
    let out = run_prints(
        r#"
        fun main() {
            val v1 = java.util.UUID.fromString("f81d4fae-7dec-11d0-a765-00a0c91e6bf6")
            val v4 = java.util.UUID.randomUUID()
            println(v1.version())
            println(v4.version())
            println(v1 != v4)
        }
    "#,
    );
    assert_eq!(out, &["1", "4", "true"]);
}
