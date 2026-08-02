// vybe-test: kotlin/type_cast_edges/test_function_type_is_check
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun op(v: Int): Int = v + 1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = ::op
            __check((value is (Int) -> Int).toString(), "true")
            val f = value as? (Int) -> Int
            __check((f?.invoke(3) ?: -1).toString(), "4")
        }
