// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_casts
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = 255u
            val b = 255u.toUByte()
            val s = 65535u.toUShort()
            val l = 1024uL
            __check((u.toByte().toString()).toString(), "-1")
            __check((b.toInt()).toString(), "255")
            __check((s.toInt()).toString(), "65535")
            __check((l.toInt()).toString(), "1024")
        }
