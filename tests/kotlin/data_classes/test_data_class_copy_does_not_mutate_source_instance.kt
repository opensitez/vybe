// vybe-test: kotlin/data_classes/test_data_class_copy_does_not_mutate_source_instance
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Pair(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = Pair(1, 2)
            val copy = original.copy(y = 9)
            __check((original.y).toString(), "2")
            __check((copy.y).toString(), "9")
        }
