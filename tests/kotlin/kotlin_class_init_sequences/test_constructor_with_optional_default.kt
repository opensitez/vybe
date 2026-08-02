// vybe-test: kotlin/kotlin_class_init_sequences/test_constructor_with_optional_default
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Entry(val a: Int = 1, val b: Int = 2) {
            val sum = a + b
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Entry(3).sum).toString(), "5")
            __check((Entry().sum).toString(), "3")
        }
