// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_indexing
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = uintArrayOf(10u, 20u, 30u)
            val b = ubyteArrayOf(11u, 12u)
            val s = ushortArrayOf(101u, 102u)
            val l = ulongArrayOf(1000uL, 2000uL)
            __check((u[1].toString()).toString(), "20")
            __check((b[0].toString()).toString(), "11")
            __check((s[1].toString()).toString(), "102")
            __check((l[1].toString()).toString(), "2000")
        }
