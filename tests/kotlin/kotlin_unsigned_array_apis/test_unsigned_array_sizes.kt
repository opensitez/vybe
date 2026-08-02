// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_sizes
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = uintArrayOf(1u, 2u, 3u)
            val b = ubyteArrayOf(4u, 5u, 6u)
            val s = ushortArrayOf(7u, 8u, 9u)
            val l = ulongArrayOf(10uL, 11uL, 12uL)
            __check((u.size).toString(), "3")
            __check((b.size).toString(), "3")
            __check((s.size).toString(), "3")
            __check((l.size).toString(), "3")
        }
