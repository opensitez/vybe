// vybe-test: kotlin/extension_functions/test_extension_function_for_nullable_interface_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

interface Labelable {
            fun label(): String
        }

        fun Labelable?.labelOrFallback(): String {
            return this?.label() ?: "missing"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val missing: Labelable? = null
            val item: Labelable = object : Labelable {
                override fun label(): String = "ok"
            }
            __check((item.labelOrFallback()).toString(), "ok")
            __check((missing.labelOrFallback()).toString(), "missing")
        }
