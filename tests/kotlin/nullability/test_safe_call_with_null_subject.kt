// vybe-test: kotlin/nullability/test_safe_call_with_null_subject
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Item {
            fun value(): Int = 7
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Item? = null
            __check((item?.value() ?: -1).toString(), "-1")
        }
