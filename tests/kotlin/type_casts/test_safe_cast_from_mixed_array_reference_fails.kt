// vybe-test: kotlin/type_casts/test_safe_cast_from_mixed_array_reference_fails
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

class Holder {
            val payload: Any = arrayOf(1, "x")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Holder().payload
            val casted = value as? Array<Int>
            __check((casted == null).toString(), "true")
        }
