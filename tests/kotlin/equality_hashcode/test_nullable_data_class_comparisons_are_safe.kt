// vybe-test: kotlin/equality_hashcode/test_nullable_data_class_comparisons_are_safe
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Box(val value: Int?)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty: Box? = null
            val left: Box? = Box(null)
            val right: Box? = Box(null)
            __check((empty == null).toString(), "true")
            __check((left == right).toString(), "true")
            __check((left == null).toString(), "false")
            __check((left === right).toString(), "false")
        }
