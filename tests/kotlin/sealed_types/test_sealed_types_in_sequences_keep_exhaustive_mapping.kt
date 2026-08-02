// vybe-test: kotlin/sealed_types/test_sealed_types_in_sequences_keep_exhaustive_mapping
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Action {
            class Append(val value: String) : Action()
            class Multiply(val value: Int) : Action()
        }

        fun emit(actions: List<Action>): String {
            return actions.joinToString(",") {
                when (it) {
                    is Action.Append -> it.value
                    is Action.Multiply -> "x" + it.value.toString()
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val actions = listOf(
                Action.Append("a"),
                Action.Multiply(2),
                Action.Append("b"),
            )
            __check((emit(actions)).toString(), "a,x2,b")
        }
