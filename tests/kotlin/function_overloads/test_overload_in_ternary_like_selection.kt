// vybe-test: kotlin/function_overloads/test_overload_in_ternary_like_selection
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun convert(v: Int): Int = v
        fun convert(v: String): Int = v.length
        fun pick(flag: Boolean, value: Int): Int = if (flag) convert(value) else convert(value.toString())
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(true, 3)).toString(), "3")
            __check((pick(false, 3)).toString(), "1")
        }
