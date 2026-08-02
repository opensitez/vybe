// vybe-test: kotlin/collections_set/test_set_average_on_empty_handled_by_exception
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val avg = setOf<Int>().average()
            __check((avg.isNaN()).toString(), "true")
        }
