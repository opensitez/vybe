// vybe-test: kotlin/kotlin_set_apis/test_set_any_none_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 4)
            __check((set.any { it > 3 }).toString(), "true")
            __check((set.none { it < 0 }).toString(), "true")
            __check((set.all { it < 10 }).toString(), "true")
        }
