use super::helpers::run_python;

// textwrap depth already done; uuid — uuid1/3/4/5, UUID fields, NAMESPACE_*, bytes/int

#[test]
fn test_uuid_uuid4_is_version_4() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(u.version)
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_uuid_uuid4_string_format() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
s = str(u)
parts = s.split("-")
print(len(parts))
print([len(p) for p in parts])
"#,
    );
    assert_eq!(out, vec!["5", "[8, 4, 4, 4, 12]"]);
}

#[test]
fn test_uuid_uuid4_uniqueness() {
    let out = run_python(
        r#"
import uuid
ids = {str(uuid.uuid4()) for _ in range(100)}
print(len(ids))
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_uuid_uuid4_bytes_length() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(len(u.bytes))
"#,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn test_uuid_uuid4_int_is_128_bit() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(0 <= u.int < 2**128)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_uuid4_fields_tuple_length() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(len(u.fields))
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_uuid_uuid3_deterministic_with_dns() {
    let out = run_python(
        r#"
import uuid
u1 = uuid.uuid3(uuid.NAMESPACE_DNS, "python.org")
u2 = uuid.uuid3(uuid.NAMESPACE_DNS, "python.org")
print(u1 == u2)
print(u1.version)
"#,
    );
    assert_eq!(out, vec!["True", "3"]);
}

#[test]
fn test_uuid_uuid5_deterministic_with_url() {
    let out = run_python(
        r#"
import uuid
u1 = uuid.uuid5(uuid.NAMESPACE_URL, "https://example.com")
u2 = uuid.uuid5(uuid.NAMESPACE_URL, "https://example.com")
print(u1 == u2)
print(u1.version)
"#,
    );
    assert_eq!(out, vec!["True", "5"]);
}

#[test]
fn test_uuid_uuid3_vs_uuid5_differ() {
    let out = run_python(
        r#"
import uuid
u3 = uuid.uuid3(uuid.NAMESPACE_DNS, "example.com")
u5 = uuid.uuid5(uuid.NAMESPACE_DNS, "example.com")
print(u3 != u5)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_namespace_dns_is_uuid() {
    let out = run_python(
        r#"
import uuid
print(isinstance(uuid.NAMESPACE_DNS, uuid.UUID))
print(str(uuid.NAMESPACE_DNS))
"#,
    );
    assert_eq!(out, vec!["True", "6ba7b810-9dad-11d1-80b4-00c04fd430c8"]);
}

#[test]
fn test_uuid_namespace_url_is_uuid() {
    let out = run_python(
        r#"
import uuid
print(str(uuid.NAMESPACE_URL))
"#,
    );
    assert_eq!(out, vec!["6ba7b811-9dad-11d1-80b4-00c04fd430c8"]);
}

#[test]
fn test_uuid_from_string_round_trip() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
u2 = uuid.UUID(str(u))
print(u == u2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_from_bytes_round_trip() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
u2 = uuid.UUID(bytes=u.bytes)
print(u == u2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_from_int_round_trip() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
u2 = uuid.UUID(int=u.int)
print(u == u2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_variant_rfc_4122() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(u.variant == uuid.RFC_4122)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_time_low_field_is_int() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(isinstance(u.time_low, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_clock_seq_field_range() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
cs = u.clock_seq
print(0 <= cs < 2**14)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_hex_is_32_chars_no_dashes() {
    let out = run_python(
        r#"
import uuid
u = uuid.uuid4()
print(len(u.hex))
print("-" not in u.hex)
"#,
    );
    assert_eq!(out, vec!["32", "True"]);
}

#[test]
fn test_uuid_uuid3_different_names_differ() {
    let out = run_python(
        r#"
import uuid
u1 = uuid.uuid3(uuid.NAMESPACE_DNS, "alpha.example.com")
u2 = uuid.uuid3(uuid.NAMESPACE_DNS, "beta.example.com")
print(u1 != u2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_uuid_invalid_string_raises_value_error() {
    let out = run_python(
        r#"
import uuid
try:
    uuid.UUID("not-a-valid-uuid")
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["ValueError"]);
}
