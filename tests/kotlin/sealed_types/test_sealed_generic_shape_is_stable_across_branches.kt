// vybe-test: kotlin/sealed_types/test_sealed_generic_shape_is_stable_across_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Value {
            class Text(val value: String) : Value()
            class Count(val value: Int) : Value()
        }

        fun normalize(value: Value): String {
            return when (value) {
                is Value.Text -> value.value
                is Value.Count -> value.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list: List<Value> = listOf(Value.Text("x"), Value.Count(3))
            __check((normalize(list[0])).toString(), "x")
            __check((normalize(list[1])).toString(), "3")
        }
