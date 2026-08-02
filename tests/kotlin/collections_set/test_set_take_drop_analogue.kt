// vybe-test: kotlin/collections_set/test_set_take_drop_analogue
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val firstTwo = values.take(2)
            val dropped = values.drop(2)
            __check((firstTwo.size).toString(), "2")
            __check((dropped.size).toString(), "3")
            __check((firstTwo.contains(1)).toString(), "true")
            __check((dropped.contains(5)).toString(), "true")
        }
