// vybe-test: kotlin/local_classes/test_local_class_array_like
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Ints {
                val items = intArrayOf(1, 2, 3)
            }
            __check((Ints().items.sum()).toString(), "6")
        }
