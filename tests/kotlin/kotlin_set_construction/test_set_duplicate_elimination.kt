// vybe-test: kotlin/kotlin_set_construction/test_set_duplicate_elimination
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_construction.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = setOf(1, 1, 2, 3)
            __check((s.size).toString(), "3")
            __check((s.contains(2)).toString(), "true")
        }
