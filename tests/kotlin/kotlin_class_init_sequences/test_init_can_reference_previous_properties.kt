// vybe-test: kotlin/kotlin_class_init_sequences/test_init_can_reference_previous_properties
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Product {
            val base = 3
            val doubled = base * 2
            init {
                __check((doubled).toString(), "6")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Product().base).toString(), "3")
        }
