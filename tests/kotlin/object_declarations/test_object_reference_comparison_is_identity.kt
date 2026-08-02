// vybe-test: kotlin/object_declarations/test_object_reference_comparison_is_identity
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Holder {
            val value = 3
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder === Holder).toString(), "true")
            __check((Holder === object : Any() { val value = 3 }).toString(), "false")
        }
