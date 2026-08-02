// vybe-test: kotlin/kotlin_uuid/test_uuid_timestamped_time_ordering_for_v1_or_not
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v1 = java.util.UUID.fromString("f81d4fae-7dec-11d0-a765-00a0c91e6bf6")
            val v4 = java.util.UUID.randomUUID()
            __check((v1.version()).toString(), "1")
            __check((v4.version()).toString(), "4")
            __check((v1 != v4).toString(), "true")
        }
