// vybe-test: kotlin/sealed_types/test_sealed_class_with_object_leaf
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Option {
            object Empty : Option()
            class Value(val value: Int) : Option()
        }

        fun label(value: Option): String {
            return when (value) {
                is Option.Empty -> "empty"
                is Option.Value -> value.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(Option.Empty)).toString(), "empty")
            __check((label(Option.Value(5))).toString(), "5")
        }
