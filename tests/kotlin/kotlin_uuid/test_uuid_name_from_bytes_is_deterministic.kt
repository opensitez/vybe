// vybe-test: kotlin/kotlin_uuid/test_uuid_name_from_bytes_is_deterministic
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = "kotlin".toByteArray()
            val a = java.util.UUID.nameUUIDFromBytes(bytes)
            val b = java.util.UUID.nameUUIDFromBytes(bytes)
            __check((a).toString(), "7f1d0d8e-2f1f-3138-92e7-4b3d8f5ef2d6")
            __check((a == b).toString(), "true")
            __check((a.version()).toString(), "3")
        }
