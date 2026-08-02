// vybe-test: kotlin/kotlin_nothing_type/test_nothing_type_in_union_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_nothing_type.rs

fun boom(): Nothing = throw Exception("nope")

        fun valueOrBoom(v: Int): Int {
            return if (v > 0) v else boom()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((valueOrBoom(2)).toString(), "2")
        }
