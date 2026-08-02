// vybe-test: kotlin/kotlin_class_init_sequences/test_property_initialized_from_constructor_parameter_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class User(name: String) {
            val upper = name.uppercase()
            init { __check((upper).toString(), "AB") }
            val size = name.length
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((User("ab").size).toString(), "2")
        }
