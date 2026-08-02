// vybe-test: kotlin/member_references/test_top_level_reference_as_higher_order_value
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun identity(v: Int) = v + 10
        fun apply(value: Int, fn: (Int) -> Int): Int = fn(value)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(3, ::identity)).toString(), "13")
        }
