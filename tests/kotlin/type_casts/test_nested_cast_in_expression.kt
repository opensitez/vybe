// vybe-test: kotlin/type_casts/test_nested_cast_in_expression
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun convert(input: Any): Int {
            return if (input is Int) {
                input as Int
            } else {
                0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((convert(9)).toString(), "9")
            __check((convert("bad")).toString(), "0")
        }
