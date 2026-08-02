// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = uintArrayOf()
            __check((u.size).toString(), "0")
            __check((u.isEmpty().toString()).toString(), "true")
            __check((ubyteArrayOf().size).toString(), "0")
            __check((ushortArrayOf().size).toString(), "0")
            __check((ulongArrayOf().size).toString(), "0")
        }
