// vybe-test: kotlin/generics/test_generic_star_projection_readonly
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> consumeUnknown(values: Array<out T>): Int {
            if (values.size == 0) {
                return 0
            }
            return values.size
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items: Array<Int> = arrayOf(1, 2, 3)
            __check((consumeUnknown(items)).toString(), "3")
            __check((consumeUnknown(arrayOf<String>())).toString(), "0")
        }
