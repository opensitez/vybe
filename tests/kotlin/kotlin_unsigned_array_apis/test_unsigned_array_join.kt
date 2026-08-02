// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_join
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
            __check((u.joinToString("|")).toString(), "1|2|3")
            __check((b.joinToString("|")).toString(), "4|5|6")
        }
