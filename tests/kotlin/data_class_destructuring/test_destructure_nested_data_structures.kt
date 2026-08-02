// vybe-test: kotlin/data_class_destructuring/test_destructure_nested_data_structures
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Left(val value: Int)
        data class Right(val other: Left, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Right(Left(7), "ok")
            val (left, label) = item
            val (value) = left
            __check((value).toString(), "7")
            __check((label).toString(), "ok")
        }
