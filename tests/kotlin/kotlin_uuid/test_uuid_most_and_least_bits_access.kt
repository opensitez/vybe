// vybe-test: kotlin/kotlin_uuid/test_uuid_most_and_least_bits_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val id = java.util.UUID.fromString("00000000-0000-0000-0000-000000000001")
            __check((id.mostSignificantBits).toString(), "0")
            __check((id.leastSignificantBits).toString(), "1")
            __check((id.toString()).toString(), "00000000-0000-0000-0000-000000000001")
        }
