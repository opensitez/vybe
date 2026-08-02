// vybe-test: kotlin/data_class_destructuring/test_destructure_into_varargs_works_as_tuple_style
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class TripleValue(val a: Int, val b: Int, val c: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (a, b, c) = TripleValue(1, 2, 3)
            __check((a + b + c).toString(), "6")
        }
