// vybe-test: kotlin/kotlin_set_apis/test_set_singleton
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(7)
            __check((set.size).toString(), "1")
            __check((set.first()).toString(), "7")
            __check((set.last()).toString(), "7")
        }
