// vybe-test: kotlin/data_class_destructuring/test_destructure_to_existing_vars
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Holder(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left
            val right
            var out = ""
            run {
                val source = Holder(9, 10)
                val (x, y) = source
                out = "${'$'}x,${'$'}y"
            }
            __check((out).toString(), "9,10")
        }
