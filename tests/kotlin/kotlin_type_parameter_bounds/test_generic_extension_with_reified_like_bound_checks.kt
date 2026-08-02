// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_extension_with_reified_like_bound_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

class Holder<T : Any>(val value: T)

        fun <T : Any> describe(value: T): String {
            return value::class.simpleName ?: ""
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Holder(1))).toString(), "Holder")
            __check((describe("x").length).toString(), "1")
        }
