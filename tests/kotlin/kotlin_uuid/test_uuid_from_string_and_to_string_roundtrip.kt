// vybe-test: kotlin/kotlin_uuid/test_uuid_from_string_and_to_string_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = "123e4567-e89b-12d3-a456-426614174000"
            val id = java.util.UUID.fromString(source)
            __check((id.toString()).toString(), "123e4567-e89b-12d3-a456-426614174000")
            __check((id.variant()).toString(), "2")
            __check((id.version()).toString(), "1")
        }
