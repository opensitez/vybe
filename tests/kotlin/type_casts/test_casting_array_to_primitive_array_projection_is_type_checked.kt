// vybe-test: kotlin/type_casts/test_casting_array_to_primitive_array_projection_is_type_checked
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: Any = arrayOf(5, 6, 7)
            val primitive = values as? IntArray
            val boxed = values as? Array<Int>
            __check((primitive == null).toString(), "true")
            __check((boxed != null).toString(), "true")
            __check((boxed?.size).toString(), "3")
        }
