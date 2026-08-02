// vybe-test: kotlin/kotlin_property_accessors_advanced/test_lazy_property_initializer
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Expensive {
            val value by lazy {
                3 + 4
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val e = Expensive()
            __check((e.value).toString(), "7")
            __check((e.value).toString(), "7")
        }
