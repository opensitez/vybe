// vybe-test: kotlin/kotlin_set_apis/test_set_subtract
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2, 3, 4)
            val b = setOf(2, 4)
            val c = a - b
            __check((c.joinToString(",")).toString(), "1,3")
        }
