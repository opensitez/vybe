// vybe-test: kotlin/function_types/test_function_type_alias_basic
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

typealias IntOp = (Int) -> Int
        fun transform(v: Int, op: IntOp): Int = op(v)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val square: IntOp = { it * it }
            __check((transform(5, square)).toString(), "25")
        }
