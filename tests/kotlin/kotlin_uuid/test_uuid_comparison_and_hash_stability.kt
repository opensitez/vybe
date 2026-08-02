// vybe-test: kotlin/kotlin_uuid/test_uuid_comparison_and_hash_stability
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            val b = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            __check((a == b).toString(), "true")
            __check((a.hashCode() == b.hashCode()).toString(), "true")
        }
